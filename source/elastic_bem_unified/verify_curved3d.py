"""Optional fine tri6 sphere refinement against uniform affine elasticity."""
import json,time
from pathlib import Path
import numpy as np
from threadpoolctl import threadpool_limits
from elasticbem import Material,Dense,fields
from verify import sphere,strain_for

def main():
    with threadpool_limits(limits=1):
        start=time.perf_counter();mesh=sphere(2);mat=Material();strain=strain_for(mat);u=mesh.x@strain.T;exact=mesh.normal@mat.stress(strain).T;op=Dense(mesh,mat);t=np.linalg.solve(op.Gmatrix,op.Hmatrix@u.ravel()).reshape(-1,3);uv,sv,reports=fields([[.1,.2,.1]],mesh,mat,u,t);error=np.max(abs(sv[0]-mat.stress(strain)));assert error<.001 and reports[0]['converged']
        result=dict(elements=mesh.ne,interior_stress_error_kPa=float(error),traction_error_kPa=float(np.max(abs(t-exact))),seconds=time.perf_counter()-start)
    Path('validation').mkdir(exist_ok=True);Path('validation/curved_tri6_fine.json').write_text(json.dumps(result,indent=2),encoding='utf-8');print(result)
if __name__=='__main__':main()
