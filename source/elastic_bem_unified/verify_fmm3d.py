import json,time
from pathlib import Path
import numpy as np
from threadpoolctl import threadpool_limits
from generate_examples import box_mesh
from elasticbem import Material,Mesh,Dense,FMM
with threadpool_limits(limits=1):
 v,e,_=box_mesh(3,4); m=Mesh(v,e,3,2,integration=dict(boundary_order=8,boundary_de_level=0,self_order=16)); mat=Material(); rng=np.random.default_rng(5); u=rng.normal(size=m.x.shape); t=rng.normal(size=m.x.shape)*mat.E
 start=time.perf_counter(); dense=Dense(m,mat); exact=dense.apply(t,u); rows=[]; print('dense',time.perf_counter()-start,flush=True)
 for p in [4,6]:
  start=time.perf_counter(); f=FMM(m,mat,p=p,leaf=12,theta=.7); y=f.apply(t,u); error=np.linalg.norm(y-exact)/np.linalg.norm(exact); assert len(f.far)>0
  rows.append(dict(p=p,far_pairs=len(f.far),relative_operator_error=float(error),seconds=time.perf_counter()-start)); print(rows[-1],flush=True)
 assert rows[-1]['relative_operator_error']<rows[0]['relative_operator_error'] and rows[-1]['relative_operator_error']<1e-6
 Path('validation/fmm3d.json').write_text(json.dumps(rows,indent=2))
