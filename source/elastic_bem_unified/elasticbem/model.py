"""Shared ND multi-region fully bonded solve. No half-space branch."""
import re,time
import numpy as np
from scipy.sparse.linalg import LinearOperator,gmres
from .meshes import Mesh
from .material import Material
from .operators import choose_operator,available_memory

def rigid_matrix(x,d):
    if d==2: return np.array([[1.,0.,-x[1]],[0.,1.,x[0]]])
    skew=np.array([[0.,-x[2],x[1]],[x[2],0.,-x[0]],[-x[1],x[0],0.]])
    return np.c_[np.eye(3),-skew]

def build(model):
    d=model.get('dimension')
    if d not in (2,3): raise ValueError('Specify dimension: 2 or 3.')
    if d==3 and 'state' in model and model['state']!='three_dimensional': raise ValueError('3D state is three_dimensional; plane states apply only to 2D.')
    if not isinstance(model.get('regions'),list) or not model['regions']: raise ValueError('At least one finite region required.')
    names={}; regions=[]; settings=dict(model.get('solver',{}))
    for spec in model['regions']:
        if spec.get('kind','finite')!='finite' or 'reference_x' in spec: raise ValueError('Only closed finite regions are supported; infinite-boundary input has been removed.')
        name=spec['name']
        if not isinstance(name,str) or not re.fullmatch(r'[A-Za-z0-9_-]+',name) or name in names: raise ValueError('Unique ASCII region name required.')
        order=spec.get('element_order',model.get('element_order',2)); mesh=Mesh(spec['vertices'],spec['elements'],d,order,model.get('integration')); k=mesh.nfield
        material=spec['material']; mat=Material(material['E'],material['nu'],d,model.get('state','plane_strain'),material.get('log_reference',mesh.scale))
        mask=np.zeros_like(mesh.x,bool); values=np.zeros_like(mesh.x); assigned=np.zeros(mesh.ne,bool)
        for bc in spec.get('boundary_conditions',[]):
            ee=np.asarray(bc['elements'])
            if ee.ndim!=1 or not np.issubdtype(ee.dtype,np.integer) or np.any(ee<0) or np.any(ee>=mesh.ne) or np.any(assigned[ee]) or len(set(ee.tolist()))!=len(ee): raise ValueError('Invalid/duplicate BC element indices.')
            assigned[ee]=True; kinds=bc['type']; vv=np.asarray(bc['values'],float)
            if len(kinds)!=d or any(v not in ('u','t') for v in kinds) or vv.shape not in ((d,),(k,d)) or not np.isfinite(vv).all(): raise ValueError('BC type/values must match dimension and field node count.')
            for e in ee: mask[k*e:k*(e+1)]=np.array(kinds)=='u'; values[k*e:k*(e+1)]=vv
        names[name]=len(regions); regions.append(dict(name=name,mesh=mesh,mat=mat,mask=mask,values=values,spec=spec,assigned=assigned))
    pairs=[]; paired=set(); graph=[set() for _ in regions]
    for interface in model.get('interfaces',[]):
        aidx=names[interface['region_a']]; bidx=names[interface['region_b']]
        if aidx==bidx: raise ValueError('Interface requires two different regions.')
        a=regions[aidx]; b=regions[bidx]; ma=a['mesh']; mb=b['mesh']; graph[aidx].add(bidx); graph[bidx].add(aidx)
        if ma.order!=mb.order: raise ValueError('Interface field orders and geometry must conform.')
        ea=interface['elements_a']; eb=interface['elements_b']; tol=1e-9*max(ma.scale,mb.scale); k=ma.nfield
        if not ea or len(ea)!=len(eb): raise ValueError('Matching interface element counts required.')
        for e,f in zip(ea,eb):
            if not isinstance(e,int) or not isinstance(f,int) or not 0<=e<ma.ne or not 0<=f<mb.ne: raise ValueError('Invalid interface element.')
            if a['assigned'][e] or b['assigned'][f]: raise ValueError('Do not prescribe exterior BC on an interface.')
            ca=ma.curves[e]; cb=mb.curves[f]
            if d==2:
                geom_good=min(np.max(abs(ca-cb)),np.max(abs(ca-cb[::-1])))<=tol
            else:
                geom_good=np.max(np.min(np.linalg.norm(ca[:,None]-cb[None],axis=-1),axis=1))<=tol and np.max(np.min(np.linalg.norm(ca[:3,None]-cb[None,:3],axis=-1),axis=1))<=tol
            if not geom_good: raise ValueError('Interface geometry including midside nodes must match.')
            for i in range(k*e,k*(e+1)):
                j=k*f+int(np.argmin(np.linalg.norm(mb.x[k*f:k*(f+1)]-ma.x[i],axis=1)))
                if np.linalg.norm(ma.x[i]-mb.x[j])>tol or np.linalg.norm(ma.normal[i]+mb.normal[j])>1e-8: raise ValueError('Interface field points and opposite normals must match.')
                if (aidx,i) in paired or (bidx,j) in paired: raise ValueError('Multiply assigned interface field node.')
                paired.update([(aidx,i),(bidx,j)]); pairs.append((aidx,i,bidx,j))
    seen=set()
    for root in range(len(regions)):
        if root in seen: continue
        component={root}; todo=[root]
        while todo:
            for k in graph[todo.pop()]:
                if k not in component: component.add(k); todo.append(k)
        seen|=component; xyz=np.concatenate([regions[k]['mesh'].x for k in component]); center=xyz.mean(axis=0); scale=max(np.ptp(xyz,axis=0).max(),1e-20); constraints=[]
        for k in component:
            r=regions[k]
            for i,j in np.argwhere(r['mask']): constraints.append(rigid_matrix((r['mesh'].x[i]-center)/scale,d)[j])
        count=3 if d==2 else 6
        if len(constraints)<count or np.linalg.matrix_rank(constraints)<count: raise ValueError(f'Connected component has unresolved {count} rigid-body modes.')
    configured_budget=settings.pop('memory_mb',None)
    if configured_budget is not None and (not np.isfinite(configured_budget) or configured_budget<=0): raise ValueError('Positive finite model memory budget required.')
    total_budget=min(.6*available_memory()/1024**2,configured_budget if configured_budget is not None else float('inf')); remaining=total_budget
    for r in regions:
        r['op'],r['selection']=choose_operator(r['mesh'],r['mat'],memory_mb=remaining,**settings); remaining-=r['selection'][r['selection']['selected']+'_estimated_MB']; r['selection']['model_memory_budget_MB']=total_budget; r['selection']['remaining_model_memory_budget_MB']=remaining
    return regions,pairs,paired

def solve_model(model):
    start=time.perf_counter(); regions,pairs,paired=build(model); d=model['dimension']; offsets=np.r_[0,np.cumsum([d*len(r['mesh'].x) for r in regions])]; n=int(offsets[-1]); groups=[]; cursor=0
    for k,r in enumerate(regions):
        for i in range(len(r['mesh'].x)):
            if (k,i) not in paired: groups.append(('external',[(k,i)],cursor,r['mat'].E/r['mesh'].scale)); cursor+=d
    for a,i,b,j in pairs:
        scale=np.sqrt(regions[a]['mat'].E*regions[b]['mat'].E)/max(regions[a]['mesh'].scale,regions[b]['mesh'].scale); groups.append(('interface',[(a,i),(b,j)],cursor,scale)); cursor+=2*d
    if cursor!=n: raise ValueError('Unknown/equation count mismatch.')
    u0=[np.where(r['mask'],r['values'],0.) for r in regions]; t0=[np.where(r['mask'],0.,r['values']) for r in regions]
    for a,i,b,j in pairs: u0[a][i]=u0[b][j]=0.; t0[a][i]=t0[b][j]=0.
    def decode(z):
        uu=[np.zeros_like(v) for v in u0]; tt=[np.zeros_like(v) for v in t0]
        for kind,nodes,c,scale in groups:
            if kind=='external':
                k,i=nodes[0]; mask=regions[k]['mask'][i]; tt[k][i]=np.where(mask,z[c:c+d]*scale,0.); uu[k][i]=np.where(mask,0.,z[c:c+d])
            else:
                (a,i),(b,j)=nodes; uu[a][i]=uu[b][j]=z[c:c+d]; tt[a][i]=z[c+d:c+2*d]*scale; tt[b][j]=-tt[a][i]
        return uu,tt
    def apply(uu,tt): return np.concatenate([r['op'].apply(-t,u).ravel() for r,u,t in zip(regions,uu,tt)])
    rhs=-apply(u0,t0); A=LinearOperator((n,n),matvec=lambda z:apply(*decode(z)),dtype=float); inverses=[]
    for kind,nodes,c,scale in groups:
        rows=np.concatenate([np.arange(offsets[k]+d*i,offsets[k]+d*(i+1)) for k,i in nodes])
        if kind=='external':
            k,i=nodes[0]; op=regions[k]['op']; B=np.where(regions[k]['mask'][i][None],-op.diagG[i]*scale,op.diagH[i])
        else:
            (a,i),(b,j)=nodes; oa=regions[a]['op']; ob=regions[b]['op']; B=np.block([[oa.diagH[i],-oa.diagG[i]*scale],[ob.diagH[j],ob.diagG[j]*scale]])
        inverses.append((rows,c,np.linalg.pinv(B,rcond=1e-12)))
    def pre(v):
        out=np.zeros(n)
        for rows,c,inv in inverses: out[c:c+len(rows)]=inv@v[rows]
        return out
    config=model.get('linear_solver',{}); rtol=config.get('rtol',1e-9); restart=config.get('restart',100); maxiter=config.get('maxiter',200)
    if not 0<rtol<1 or restart<1 or maxiter<1: raise ValueError('Invalid GMRES controls.')
    history=[]; z,info=gmres(A,rhs,M=LinearOperator((n,n),matvec=pre,dtype=float),rtol=rtol,atol=0,restart=restart,maxiter=maxiter,callback=history.append,callback_type='pr_norm'); residual=np.linalg.norm(A@z-rhs)/max(np.linalg.norm(rhs),1e-300); good=info==0 and residual<=max(2*rtol,1e-12)
    uu,tt=decode(z); uu=[u+v for u,v in zip(uu,u0)]; tt=[t+v for t,v in zip(tt,t0)]
    report=dict(dimension=d,converged=bool(good),error_code='OK' if good else 'LINEAR_SOLVE_TOLERANCE_UNMET',gmres_info=int(info),true_relative_residual=float(residual),iterations=len(history),seconds=time.perf_counter()-start,interface_field_nodes=len(pairs),interface_displacement_jump_max=float(max([np.max(abs(uu[a][i]-uu[b][j])) for a,i,b,j in pairs] or [0.])),interface_traction_imbalance_max=float(max([np.max(abs(tt[a][i]+tt[b][j])) for a,i,b,j in pairs] or [0.])))
    return regions,uu,tt,report
