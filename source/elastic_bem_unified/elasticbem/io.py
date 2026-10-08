"""Common CSV/NPZ/VTK outputs for both dimensions; save unmet-accuracy values."""
import csv,json,sys
from pathlib import Path
import numpy as np
from .model import solve_model
from .precision import fields,QuadraturePolicy,json_safe
from .meshes import TRI_GEOM

def vtk(path,mesh,u,t):
    d=mesh.dimension; k=mesh.nfield; ng=mesh.elements.shape[1]; xyz=mesh.curves.reshape(-1,d); s=(np.array([[-1.],[0.],[1.]]) if mesh.order==2 else np.array([[-1.],[1.]])) if d==2 else (TRI_GEOM if mesh.order==2 else TRI_GEOM[:3]); N=mesh.shape(s)
    uu=np.concatenate([N@u[k*e:k*(e+1)] for e in range(mesh.ne)]); tt=np.concatenate([N@t[k*e:k*(e+1)] for e in range(mesh.ne)]); permutation=[0,2,1] if d==2 and mesh.order==2 else list(range(ng)); celltype=(21 if mesh.order==2 else 3) if d==2 else (22 if mesh.order==2 else 5)
    with open(path,'w',encoding='ascii') as f:
        f.write('# vtk DataFile Version 3.0\nUnified elasticity boundary, discontinuous fields\nASCII\nDATASET UNSTRUCTURED_GRID\n'); f.write(f'POINTS {len(xyz)} double\n')
        for p in xyz: f.write(' '.join(f'{a:.16g}' for a in np.r_[p,np.zeros(3-d)])+'\n')
        f.write(f'CELLS {mesh.ne} {(ng+1)*mesh.ne}\n')
        for e in range(mesh.ne): f.write(str(ng)+' '+' '.join(str(ng*e+i) for i in permutation)+'\n')
        f.write(f'CELL_TYPES {mesh.ne}\n'+(str(celltype)+'\n')*mesh.ne); f.write(f'POINT_DATA {len(xyz)}\n')
        for name,values in [('displacement_m',uu),('traction_kPa',tt)]:
            f.write(f'VECTORS {name} double\n')
            for p in values: f.write(' '.join(f'{a:.16g}' for a in np.r_[p,np.zeros(3-d)])+'\n')

def save(model,out):
    policy=QuadraturePolicy(**model.get('quadrature',{})); out=Path(out); out.mkdir(parents=True,exist_ok=True); regions,uu,tt,linear=solve_model(model); arrays=dict(model_json=np.array(json.dumps(model)),dimension=np.array(model['dimension'])); reports=[]; good=linear['converged']; d=model['dimension']
    for index,(r,u,t) in enumerate(zip(regions,uu,tt)):
        mesh=r['mesh']; prefix=f'region_{index}_'; arrays.update({prefix+'vertices':mesh.vertices,prefix+'elements':mesh.elements,prefix+'collocation':mesh.x,prefix+'u':u,prefix+'t':t}); axis=list('xyz'[:d])
        np.savetxt(out/(r['name']+'_boundary.csv'),np.c_[mesh.x,u,t],delimiter=',',header=','.join([a+'_m' for a in axis]+['u'+a+'_m' for a in axis]+['t'+a+'_kPa' for a in axis]),comments=''); vtk(out/(r['name']+'_boundary.vtk'),mesh,u,t)
        s,w=mesh.regular_rule(32 if d==2 else 16); force=np.zeros(d)
        for e in range(mesh.ne): force+=(w*mesh.geometry(e,s)[1])@(mesh.shape(s)@t[mesh.nfield*e:mesh.nfield*(e+1)])
        rr=dict(name=r['name'],material=r['spec']['material'],element_order=mesh.order,elements=mesh.ne,field_nodes=len(mesh.x),selection=r['selection'],operator=r['op'].stats(),resultant_traction_kN=force.tolist(),resultant_units='kN/m of out-of-plane length' if d==2 else 'kN'); points=r['spec'].get('interior_points',[])
        if points:
            uv,sv,pr=fields(points,mesh,r['mat'],u,t,policy); arrays.update({prefix+'interior_points':np.asarray(points),prefix+'interior_u':uv,prefix+'interior_stress':sv}); names=['sxx','syy','szz','sxy','syz','szx']
            with open(out/(r['name']+'_interior.csv'),'w',newline='',encoding='utf-8') as f:
                writer=csv.writer(f); writer.writerow([a+'_m' for a in axis]+['u'+a+'_m' for a in axis]+[a+'_kPa' for a in names]+['quadrature_converged','error_code','value_available'])
                for point,displacement,stress,check in zip(points,uv,sv,pr): writer.writerow([*point,*displacement,stress[0,0],stress[1,1],stress[2,2],stress[0,1],stress[1,2],stress[2,0],check['converged'],check['error_code'],check['value_available']])
            rr['interior_reports']=pr; good &=all(p['converged'] for p in pr)
            for i,p in enumerate(pr):
                if not p['converged']: print(f"ERROR: {r['name']} interior point {i}: {p['error_code']}; available values retained.",file=sys.stderr)
        reports.append(rr)
    if not linear['converged']: print('ERROR: LINEAR_SOLVE_TOLERANCE_UNMET; iterate retained.',file=sys.stderr)
    np.savez_compressed(out/'solution.npz',**arrays); report=dict(dimension=d,linear_solver=linear,regions=reports,all_requested_checks_passed=bool(good),accuracy_scope='quadrature and algebraic residual only; discretization/FMM errors require convergence studies'); (out/'report.json').write_text(json.dumps(json_safe(report),ensure_ascii=False,indent=2,allow_nan=False),encoding='utf-8'); print(json.dumps(dict(output=str(out),converged=bool(good),linear_solver=linear),ensure_ascii=False)); return 0 if good else 2
