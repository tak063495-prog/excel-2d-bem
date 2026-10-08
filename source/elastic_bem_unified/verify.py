"""Finite-domain ND regression: physics, curved tri6, bonded interfaces, output."""
import json,time,tempfile,subprocess,sys
from pathlib import Path
import numpy as np
from scipy.spatial import ConvexHull
from threadpoolctl import threadpool_limits
from elasticbem import *
from elasticbem.material import kernels,gradients
from elasticbem.io import save
from elasticbem.precision import json_safe
from generate_examples import block_model,two_material_model,box_mesh

def strain_for(mat,sigma=100.):
    if mat.dimension==3: return np.diag([sigma/mat.E,-sigma*mat.nu/mat.E,-sigma*mat.nu/mat.E])
    if mat.state=='plane_stress': return np.diag([sigma/mat.E,-sigma*mat.nu/mat.E])
    return np.diag([(1-mat.nu**2)*sigma/mat.E,-mat.nu*(1+mat.nu)*sigma/mat.E])

def sphere(level):
    v=[np.array(x,float) for x in [[1,0,0],[-1,0,0],[0,1,0],[0,-1,0],[0,0,1],[0,0,-1]]]; tri=ConvexHull(v).simplices.tolist()
    for e,t in enumerate(tri):
        p=np.array(v)[t]
        if np.cross(p[1]-p[0],p[2]-p[0])@p.mean(axis=0)<0: tri[e]=[t[0],t[2],t[1]]
    for _ in range(level):
        lookup={}; fine=[]
        def midpoint(a,b):
            key=tuple(sorted([a,b]))
            if key not in lookup:
                x=v[a]+v[b]; x/=np.linalg.norm(x); lookup[key]=len(v); v.append(x)
            return lookup[key]
        for a,b,c in tri:
            ab,bc,ca=midpoint(a,b),midpoint(b,c),midpoint(c,a); fine.extend([[a,ab,ca],[ab,b,bc],[ca,bc,c],[ab,bc,ca]])
        tri=fine
    lookup={}; elems=[]
    for a,b,c in tri:
        ids=[]
        for i,j in [(a,b),(b,c),(c,a)]:
            key=tuple(sorted([i,j]))
            if key not in lookup:
                x=v[i]+v[j]; x/=np.linalg.norm(x); lookup[key]=len(v); v.append(x)
            ids.append(lookup[key])
        elems.append([a,b,c,*ids])
    return Mesh(v,elems,3,2)

def run():
    result={}; start=time.perf_counter(); rng=np.random.default_rng(42)
    for d,state in [(2,'plane_strain'),(2,'plane_stress'),(3,'three_dimensional')]:
        mat=Material(dimension=d,state=state); x=np.ones(d)*-.2; y=rng.uniform(.4,1,(4,d)); n=rng.normal(size=(4,d));n/=np.linalg.norm(n,axis=1)[:,None]; dU,dT=gradients(x,y,n,mat); error=0.
        for j in range(d):
            delta=np.eye(d)[j]*1e-6; plus=kernels((x+delta)[None],y,mat);minus=kernels((x-delta)[None],y,mat);du=(plus[0][0]-minus[0][0])/2e-6;dt=np.einsum('qijk,qk->qij',(plus[1][0]-minus[1][0])/2e-6,n);error=max(error,np.linalg.norm(du-dU[:,:,:,j])/np.linalg.norm(du),np.linalg.norm(dt-dT[:,:,:,j])/np.linalg.norm(dt))
        assert error<1e-7;result[f'{d}d_{state}_kernel_gradient_relative_error']=error
        for name,fn in [('block',block_model),('two_material',two_material_model)]:
            m=fn(d,state,n=2 if d==2 else 1);m['solver']={'backend':'dense'};rr,uu,tt,report=solve_model(m);assert report['converged'];uerror=0.;terror=0.;serror=0.;uxoffset=0.
            for region,u,t in zip(rr,uu,tt):
                mesh=region['mesh'];mat=region['mat'];strain=strain_for(mat);ue=mesh.x@strain.T;ue[:,0]=(mesh.x[:,0]-mesh.vertices[:,0].min())*strain[0,0]+uxoffset;uxoffset+=strain[0,0]*np.ptp(mesh.vertices[:,0]);te=mesh.normal@mat.stress(strain)[:d,:d].T;uerror=max(uerror,np.max(abs(u-ue)));terror=max(terror,np.max(abs(t-te)));uv,sv,pr=fields(region['spec']['interior_points'],mesh,mat,u,t);assert all(a['converged'] for a in pr);serror=max(serror,np.max(abs(sv-mat.stress(strain))))
            assert uerror<1e-8 and terror<1e-3 and serror<1e-3;result[f'{d}d_{state}_{name}']=dict(displacement_max_error_m=uerror,traction_max_error_kPa=terror,stress_max_error_kPa=serror,solver=report)
        # FMM against dense in complete coupled solves (small 3D has all-near interactions).
        m=two_material_model(d,state,n=6 if d==2 else 1);m['solver']={'backend':'dense'};rd,ud,td,pd=solve_model(m);m['solver']={'backend':'fmm','p':8 if d==2 else 4,'theta':.6,'leaf':9 if d==2 else 24};rf,uf,tf,pf=solve_model(m);du=max(np.max(abs(a-b)) for a,b in zip(ud,uf));dt=max(np.max(abs(a-b)) for a,b in zip(td,tf));assert pf['converged'] and du<1e-8 and dt<1e-3;result[f'{d}d_{state}_coupled_fmm_dense']=dict(u_difference_m=du,traction_difference_kPa=dt,far_pairs=[len(r['op'].far) for r in rf])
        # Constant elements remain available through the same entry point.
        m=block_model(d,state,n=8 if d==2 else 2,order=0);m['solver']={'backend':'dense'};rr,u,t,report=solve_model(m);assert report['converged'];region=rr[0];mesh=region['mesh'];mat=region['mat'];e=strain_for(mat);ue=mesh.x@e.T;relative=np.linalg.norm(u[0]-ue)/np.linalg.norm(ue);assert relative<.1;result[f'{d}d_{state}_constant_u_relative_error']=relative
    # Curved tri6 geometry: affine prescribed displacement, solve traction independently.
    curved=[]
    for level in [0,1]:
        mesh=sphere(level);mat=Material();strain=strain_for(mat);u=mesh.x@strain.T;t_exact=mesh.normal@mat.stress(strain).T;op=Dense(mesh,mat);t=np.linalg.solve(op.Gmatrix,op.Hmatrix@u.ravel()).reshape(-1,3);uv,sv,rp=fields([[.1,.2,.1]],mesh,mat,u,t);assert rp[0]['converged'];curved.append(dict(elements=mesh.ne,interior_stress_error_kPa=np.max(abs(sv[0]-mat.stress(strain))),traction_error_kPa=np.max(abs(t-t_exact))))
    assert curved[-1]['interior_stress_error_kPa']<curved[0]['interior_stress_error_kPa'] and curved[-1]['interior_stress_error_kPa']<.01;result['curved_tri6_refinement']=curved
    # Curved line3 also reproduces affine geometry/field refinement.
    curves=[]
    for ne in [8,16]:
        ends=np.c_[np.cos(np.arange(ne)*2*np.pi/ne),np.sin(np.arange(ne)*2*np.pi/ne)]; vertices=list(ends);elems=[]
        for i in range(ne):
            theta=(i+.5)*2*np.pi/ne;vertices.append([np.cos(theta),np.sin(theta)]);elems.append([i,len(vertices)-1,(i+1)%ne])
        mesh=Mesh(vertices,elems,2,2);mat=Material(dimension=2);e=strain_for(mat);u=mesh.x@e.T;op=Dense(mesh,mat);t=np.linalg.solve(op.Gmatrix,op.Hmatrix@u.ravel()).reshape(-1,2);uv,sv,rp=fields([[.1,.2]],mesh,mat,u,t);curves.append(dict(elements=ne,interior_stress_error_kPa=np.max(abs(sv[0]-mat.stress(e)))))
    assert curves[-1]['interior_stress_error_kPa']<curves[0]['interior_stress_error_kPa'];result['curved_line3_refinement']=curves
    # DE geometry trigger and true accuracy against affine interior fields.
    v,e,_=box_mesh(3,1);mesh=Mesh(v,e,3,2);mat=Material();strain=strain_for(mat);u=mesh.x@strain.T;t=mesh.normal@mat.stress(strain).T
    uv,sv,near_reports=fields([[.5,.5,.999],[.4,.5,.99999]],mesh,mat,u,t)
    assert near_reports[0]['converged'] and all(p['de_elements']>0 for p in near_reports) and np.max(abs(sv-mat.stress(strain)))<1e-6
    result['near_boundary_3d_de']=dict(stress_max_error_kPa=np.max(abs(sv-mat.stress(strain))),reports=near_reports)
    # Memory guard, including a feasible dense solve when high-order FMM is too large.
    op,selection=choose_operator(mesh,mat,backend='dense',p=6,memory_mb=4.)
    assert selection['selected']=='dense'
    try:choose_operator(mesh,mat,memory_mb=.01)
    except MemoryError:pass
    else:raise AssertionError('Memory gate failed.')
    result['memory_gates']='passed'
    # Failure output behavior: physical quadrature tolerance unmet and GMRES iteration limit.
    for mode in ['quadrature','linear']:
        m=block_model(3);m['solver']={'backend':'dense'}
        if mode=='quadrature':m['quadrature']=dict(displacement_tol=1e-30,stress_tol=1e-30,max_points=32)
        else:m['linear_solver']=dict(restart=1,maxiter=1)
        with tempfile.TemporaryDirectory() as directory:
            code=save(m,directory);assert code==2;data=np.load(Path(directory)/'solution.npz');report=json.loads((Path(directory)/'report.json').read_text());assert np.isfinite(data['region_0_u']).all() and not report['all_requested_checks_passed'];assert (Path(directory)/'block_boundary.vtk').is_file()
            if mode=='quadrature':assert np.isfinite(data['region_0_interior_u']).all()
    result['failure_output_retention']='passed'
    for mode in ['removed_infinite','rigid','interface','dimension']:
        m=two_material_model(3) if mode=='interface' else block_model(3)
        if mode=='removed_infinite':m['regions'][0]['kind']='halfspace'
        elif mode=='rigid':m['regions'][0]['boundary_conditions']=[]
        elif mode=='dimension':m['dimension']=4
        else:m['regions'][1]['vertices'][0][0]+=.1
        try:solve_model(m)
        except ValueError:pass
        else:raise AssertionError('Invalid input accepted: '+mode)
    result['input_rejections']='passed';result['seconds']=time.perf_counter()-start;return result

if __name__=='__main__':
    with threadpool_limits(limits=1):result=run()
    Path('validation').mkdir(exist_ok=True);Path('validation/verification.json').write_text(json.dumps(json_safe(result),indent=2,allow_nan=False),encoding='utf-8');print(json.dumps(json_safe(result),indent=2,allow_nan=False))
