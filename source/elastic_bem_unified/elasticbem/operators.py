"""The same dense/FMM implementation for 2D/3D, constant/quadratic fields."""
from collections import OrderedDict
import time,os
import numpy as np
from scipy.sparse import bsr_matrix
from .material import kernels
from .tree import Chebyshev,InteractionPlan

def blocks(mesh,targets,sources,mat):
    targets=np.asarray(targets); sources=np.asarray(sources); d=mesh.dimension; k=mesh.nfield; ee=np.unique(sources//k); z,w=mesh.regular_rule(); N=mesh.shape(z)
    yy=[]; jj=[]; nn=[]
    for e in ee:
        y,j,n=mesh.geometry(e,z); yy.append(y); jj.append(j); nn.append(n)
    yy=np.array(yy); jj=np.array(jj); nn=np.array(nn); G=[]; T=[]; tile=12 if d==2 else 4
    index={int(e):a for a,e in enumerate(ee)}; cols=np.array([k*index[int(j//k)]+j%k for j in sources])
    for begin in range(0,len(targets),tile):
        rows=targets[begin:begin+tile]; U,D=kernels(mesh.x[rows],yy.reshape(-1,d),mat); U=U.reshape(len(rows),len(ee),len(z),d,d)
        tr=np.einsum('aeqijk,eqk->aeqij',D.reshape(len(rows),len(ee),len(z),d,d,d),nn)
        allG=np.einsum('eq,aeqij,ql->aelij',jj*w,U,N); allT=np.einsum('eq,aeqij,ql->aelij',jj*w,tr,N)
        for a,i in enumerate(rows):
            for pos,e in enumerate(ee):
                own=i//k==e; support=mesh.v[k*e]; bound=np.linalg.norm(np.maximum(np.maximum(support[0]-mesh.x[i],mesh.x[i]-support[1]),0.))
                if not own and bound>.2*mesh.element_length[e]: continue
                if not own and mesh.closest(e,mesh.x[i])[1]>.2*mesh.element_length[e]: continue
                s,w2=mesh.integration_rule(e,mesh.x[i],self_param=mesh.param_nodes[i%k] if own else None); y,j,n=mesh.geometry(e,s); N2=mesh.shape(s); U2,D2=kernels(mesh.x[i:i+1],y,mat); tr=np.einsum('qijk,qk->qij',D2[0],n)
                allG[a,pos]=np.einsum('q,qab,ql->lab',w2*j,U2[0],N2); field=N2.copy()
                if own: field[:,i%k]-=1
                allT[a,pos]=np.einsum('q,qab,ql->lab',w2*j,tr,field)
        G.append(allG.reshape(len(rows),-1,d,d)[:,cols]); T.append(allT.reshape(len(rows),-1,d,d)[:,cols])
    return np.concatenate(G),np.concatenate(T)

class Dense:
    def __init__(self,mesh,mat):
        self.mesh=mesh; d=mesh.dimension; n=len(mesh.x); ids=np.arange(n); self.Gmatrix=np.empty((d*n,d*n)); self.Hmatrix=np.empty_like(self.Gmatrix); self.diagG=np.empty((n,d,d)); self.diagH=np.empty_like(self.diagG); tile=12 if d==2 else 4
        for a in range(0,n,tile):
            rows=ids[a:a+tile]; G,T=blocks(mesh,rows,ids,mat); correction=-T.sum(axis=1)
            for b,i in enumerate(rows): T[b,i]+=correction[b]
            self.Gmatrix[d*a:d*(a+len(rows))]=G.transpose(0,2,1,3).reshape(d*len(rows),d*n); self.Hmatrix[d*a:d*(a+len(rows))]=T.transpose(0,2,1,3).reshape(d*len(rows),d*n); self.diagG[rows]=G[np.arange(len(rows)),rows]; self.diagH[rows]=T[np.arange(len(rows)),rows]
    def apply(self,t,u): return (self.Gmatrix@t.ravel()+self.Hmatrix@u.ravel()).reshape(self.mesh.x.shape)
    def stats(self): return dict(backend='dense',elements=self.mesh.ne,field_nodes=len(self.mesh.x),matrix_MB=(self.Gmatrix.nbytes+self.Hmatrix.nbytes)/1024**2)

class FMM:
    def __init__(self,mesh,mat,p=None,leaf=24,theta=.7,cache_mb=128,plan=None):
        d=mesh.dimension; p=p or (6 if d==2 else 4); self.mesh=mesh; self.mat=mat; self.cheb=Chebyshev(p,d); q=self.cheb.q; self.plan=plan or InteractionPlan(mesh,leaf,theta); plan=self.plan
        self.nodes=plan.nodes; self.leaves=plan.leaves; self.root=plan.root; self.far=plan.far; self.near_count=len(plan.near); self.cache=OrderedDict(); self.cache_size=0; self.cache_limit=int(cache_mb*1024**2); self.hits=0; self.misses=0
        for node in self.nodes: node.knots=node.center+node.half*self.cheb.grid
        z,w=mesh.regular_rule(max(24,2*p+4) if d==2 else max(12,2*p+4)); k=mesh.nfield
        for node in self.nodes:
            if node.children: node.transfer=[self.cheb.basis(c.knots,node.center,node.half) for c in node.children]
            else:
                node.source_t=np.zeros((q,len(node.ids))); node.source_un=np.zeros((q,len(node.ids),d))
                for e in np.unique(node.ids//k):
                    cols=np.flatnonzero(node.ids//k==e); local=node.ids[cols]%k; y,j,n=mesh.geometry(e,z); B=self.cheb.basis(y,node.center,node.half); N=mesh.shape(z)[:,local]
                    node.source_t[:,cols]=np.einsum('q,qa,ql->al',w*j,B,N); node.source_un[:,cols]=np.einsum('q,qa,ql,qk->alk',w*j,B,N,n)
                node.target=self.cheb.basis(mesh.x[node.ids],node.center,node.half)
        counts=np.zeros(len(mesh.x),int)
        for a,b in plan.near: counts[a.ids]+=len(b.ids)
        ptr=np.r_[0,np.cumsum(counts)]; cur=ptr[:-1].copy(); ids=np.empty(ptr[-1],int); gd=np.empty((ptr[-1],d,d)); td=np.empty_like(gd); self.diagG=np.zeros((len(mesh.x),d,d)); localT=np.zeros_like(self.diagG)
        # Combine all near source leaves for each target leaf. An element's six
        # source basis functions may span leaves; integrate its geometry only once.
        near_groups={}
        for a,b in plan.near: near_groups.setdefault(a.number,(a,[]))[1].append(b.ids)
        for a,source_leaves in near_groups.values():
            sources=np.sort(np.concatenate(source_leaves)); G,T=blocks(mesh,a.ids,sources,mat)
            for c,i in enumerate(a.ids):
                sl=slice(cur[i],cur[i]+len(sources)); ids[sl]=sources; gd[sl]=G[c]; td[sl]=T[c]; cur[i]+=len(sources); same=np.flatnonzero(sources==i)
                if len(same): self.diagG[i]=G[c,same[0]]; localT[i]=T[c,same[0]]
        for i in range(len(mesh.x)):
            sl=slice(ptr[i],ptr[i+1]); order=np.argsort(ids[sl]); ids[sl]=ids[sl][order]; gd[sl]=gd[sl][order]; td[sl]=td[sl][order]
        sh=(d*len(mesh.x),)*2; self.Gnear=bsr_matrix((gd,ids,ptr),shape=sh); self.Tnear=bsr_matrix((td,ids,ptr),shape=sh)
        groups={}
        for a,b in self.far: groups.setdefault(plan.key(a,b),[]).append((a,b))
        self.groups=[(v[0][0],v[0][1],np.array([a.number for a,b in v]),np.array([b.number for a,b in v])) for v in groups.values()]
        self.correction=np.zeros_like(self.diagG)
        for j in range(d):
            u=np.zeros_like(mesh.x); u[:,j]=1; self.correction[:,:,j]=-self.raw(np.zeros_like(u),u)
        self.diagH=localT+self.correction
    def m2l(self,a,b):
        key=self.plan.key(a,b)
        if key in self.cache: self.hits+=1; self.cache.move_to_end(key); return self.cache[key]
        self.misses+=1; q=self.cheb.q; d=self.mesh.dimension; U,D=kernels(a.knots,b.knots,self.mat); K=np.concatenate([U,D.reshape(q,q,d,d*d)],axis=-1).transpose(0,2,1,3).reshape(d*q,(d+d*d)*q); K=np.ascontiguousarray(K)
        if K.nbytes<=self.cache_limit:
            while self.cache and self.cache_size+K.nbytes>self.cache_limit: _,old=self.cache.popitem(last=False); self.cache_size-=old.nbytes
            self.cache[key]=K; self.cache_size+=K.nbytes
        return K
    def raw(self,t,u):
        q=self.cheb.q; d=self.mesh.dimension; nc=d+d*d; M=np.zeros((len(self.nodes),q,nc)); L=np.zeros((len(self.nodes),q,d))
        for node in self.leaves:
            M[node.number,:,:d]=node.source_t@t[node.ids]; M[node.number,:,d:]=np.einsum('ank,nj->ajk',node.source_un,u[node.ids]).reshape(q,d*d)
        for node in reversed(self.nodes):
            for c,B in zip(node.children,getattr(node,'transfer',[])): M[node.number]+=B.T@M[c.number]
        for a,b,targets,sources in self.groups:
            fields=(M[sources].reshape(len(sources),nc*q)@self.m2l(a,b).T).reshape(-1,q,d); np.add.at(L,targets,fields)
        for node in self.nodes:
            for c,B in zip(node.children,getattr(node,'transfer',[])): L[c.number]+=B@L[node.number]
        out=np.zeros_like(self.mesh.x)
        for node in self.leaves: out[node.ids]+=node.target@L[node.number]
        return out+(self.Gnear@t.ravel()+self.Tnear@u.ravel()).reshape(self.mesh.x.shape)
    def apply(self,t,u): return self.raw(t,u)+np.einsum('nij,nj->ni',self.correction,u)
    def stats(self): return dict(backend='fmm',elements=self.mesh.ne,field_nodes=len(self.mesh.x),nodes=len(self.nodes),far_pairs=len(self.far),near_pairs=self.near_count,near_matrix_MB=(self.Gnear.data.nbytes+self.Tnear.data.nbytes+self.Gnear.indices.nbytes+self.Tnear.indices.nbytes+self.Gnear.indptr.nbytes+self.Tnear.indptr.nbytes)/1024**2,m2l_cache_MB=self.cache_size/1024**2,cache_hits=self.hits,cache_misses=self.misses)

def available_memory():
    try:
        if os.name=='nt':
            import ctypes
            class Status(ctypes.Structure): _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(k,ctypes.c_ulonglong) for k in ('total','available','totalpage','availablepage','totalvirtual','availablevirtual','extended')]
            s=Status(); s.length=ctypes.sizeof(s)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(s)): return int(s.available)
        from pathlib import Path
        values={line.split(':')[0]:int(line.split()[1])*1024 for line in Path('/proc/meminfo').read_text().splitlines() if line.startswith(('MemAvailable:','MemFree:'))}; amount=values.get('MemAvailable',values.get('MemFree',512*1024**2)); limit=Path('/sys/fs/cgroup/memory.max'); used=Path('/sys/fs/cgroup/memory.current')
        if limit.exists() and used.exists() and limit.read_text().strip()!='max': amount=min(amount,max(0,int(limit.read_text())-int(used.read_text())))
        return amount
    except (OSError,ValueError): return 512*1024**2

def timing(f,repeats=3):
    f(); values=[]
    for _ in range(repeats):
        start=time.perf_counter(); f(); values.append(time.perf_counter()-start)
    return float(np.median(values))

def choose_operator(mesh,mat,backend='auto',p=None,leaf=24,theta=.7,cache_mb=128,expected_iterations=120,memory_mb=None):
    start_selection=time.perf_counter(); d=mesh.dimension; n=len(mesh.x); p=p or (6 if d==2 else 4); Chebyshev(p,d)
    if backend not in ('auto','dense','fmm') or expected_iterations<1 or cache_mb<0 or not np.isfinite([expected_iterations,cache_mb]).all(): raise ValueError('Invalid backend/calibration settings.')
    if memory_mb is not None and (memory_mb<=0 or not np.isfinite(memory_mb)): raise ValueError('Positive finite memory_mb required.')
    budget=min(.6*available_memory(),(memory_mb if memory_mb is not None else float('inf'))*1024**2); cache_mb=min(cache_mb,budget*.2/1024**2); plan=InteractionPlan(mesh,leaf,theta); q=p**d; nc=d+d*d
    dm=16*d*d*n*n+40000*n; fm=(16*d*d+24)*plan.near_entries+16*len(plan.nodes)*q*q+8*n*q*(d+1)+8*len(plan.nodes)*q*(nc+d)+min(cache_mb*1024**2,8*d*nc*q*q*plan.translation_count)+40000*n+(32*d*nc*q*q if plan.far else 0)
    if (backend=='dense' and dm>budget) or (backend=='fmm' and fm>budget) or (backend=='auto' and min(dm,fm)>budget): raise MemoryError('Estimated backend memory exceeds available model budget.')
    sample=min(n,256 if d==3 else 1600); cols=np.linspace(0,n-1,sample,dtype=int); rows=np.unique(np.linspace(0,n-1,min(2 if d==3 else 6,n),dtype=int)); start=time.perf_counter(); blocks(mesh,rows,cols,mat); dense_assembly=(time.perf_counter()-start)*n/len(rows)*n/sample
    near_groups={}
    for a,b in plan.near: near_groups.setdefault(a.number,(a,[]))[1].append(b.ids)
    sampled=list(near_groups.values()); near_seconds=0.; near_pairs=0
    for index in np.unique(np.linspace(0,len(sampled)-1,min(3,len(sampled)),dtype=int)):
        a,source_leaves=sampled[index]; sources=np.sort(np.concatenate(source_leaves)); start=time.perf_counter(); blocks(mesh,a.ids,sources,mat); near_seconds+=time.perf_counter()-start; near_pairs+=len(a.ids)*len(sources)
    near_assembly=near_seconds/max(near_pairs,1)*plan.near_entries
    rng=np.random.default_rng(31); m=min(n,256); A=rng.normal(size=(d*m,d*m)); v=rng.normal(size=d*m); dense_mv=timing(lambda:A@v+A@v)/m**2*n**2
    if fm>budget:
        prototype='fmm_memory_infeasible'; far_mv=0.; near_mv=0.; generation=0.; repeats=0
    else:
        # Execute actual hierarchy loops using bounded-memory surrogate translations.
        ghost=object.__new__(FMM); ghost.mesh=mesh; ghost.cheb=Chebyshev(p,d); ghost.nodes=plan.nodes; ghost.leaves=plan.leaves; ghost.root=plan.root
        resident=8*len(plan.nodes)*q*q+8*n*q*(d+1)+8*len(plan.nodes)*q*(nc+d); prototype='real_hierarchy_surrogate'
        K=rng.normal(size=(d*q,nc*q)) if plan.far else np.empty((0,0))
        if resident<min(128*1024**2,.15*budget):
            for node in plan.nodes:
                if node.children: node.transfer=[rng.normal(size=(q,q)) for c in node.children]
                else:
                    node.source_t=rng.normal(size=(q,len(node.ids))); node.source_un=rng.normal(size=(q,len(node.ids),d)); node.target=rng.normal(size=(len(node.ids),q))
            groups={}
            for a,b in plan.far: groups.setdefault(plan.key(a,b),[]).append((a,b))
            ghost.groups=[(v[0][0],v[0][1],np.array([a.number for a,b in v]),np.array([b.number for a,b in v])) for v in groups.values()]; ghost.Gnear=bsr_matrix((d*n,d*n)); ghost.Tnear=ghost.Gnear; ghost.m2l=lambda a,b:K; far_mv=timing(lambda:ghost.raw(mesh.x,mesh.x))
        else:
            prototype='bounded_memory_extrapolation'; source=rng.normal(size=nc*q); B=rng.normal(size=(q,q)); moments=rng.normal(size=(q,nc)); far_mv=(timing(lambda:K@source)*len(plan.far) if plan.far else 0)+2*timing(lambda:B@moments)*len(plan.nodes)+2e-7*n*q
        near_mv=3e-8*plan.near_entries; knots=ghost.cheb.grid; generation=timing(lambda:kernels(knots,knots+5,mat)) if plan.far else 0.; cache_capacity=int(cache_mb*1024**2/max(K.nbytes,1)); repeats=0 if plan.translation_count<=cache_capacity else expected_iterations*plan.translation_count
    ds=dense_assembly+expected_iterations*dense_mv; fs=near_assembly+expected_iterations*(far_mv+near_mv)+generation*(plan.translation_count+repeats)+.00002*len(plan.nodes)*q
    if fm>budget: fs=float('inf')
    choice=backend if backend!='auto' else min((time,name) for time,name,mem in [(ds,'dense',dm),(fs,'fmm',fm)] if mem<=budget)[1]
    calibration=time.perf_counter()-start_selection; start=time.perf_counter(); op=Dense(mesh,mat) if choice=='dense' else FMM(mesh,mat,p,leaf,theta,cache_mb,plan)
    return op,dict(selected=choice,dense_estimated_seconds=ds,fmm_estimated_seconds=fs,dense_estimated_MB=dm/1024**2,fmm_estimated_MB=fm/1024**2,expected_iterations=expected_iterations,selection_seconds=calibration,assembly_seconds=time.perf_counter()-start,prototype_mode=prototype,estimate_is_guarantee=False)
