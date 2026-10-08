"""Postprocess saved ND solutions without rerunning the boundary solve."""
import argparse,json,csv
from pathlib import Path
import numpy as np
from threadpoolctl import threadpool_limits
from elasticbem import Mesh,Material,QuadraturePolicy,fields
from elasticbem.precision import json_safe

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('solution');p.add_argument('points',help='CSV with x,y[,z] header');p.add_argument('--region',default=None);p.add_argument('--out',default='output_fields');p.add_argument('--u-tol',type=float,default=1e-8);p.add_argument('--stress-tol',type=float,default=1e-3);args=p.parse_args()
    with np.load(args.solution,allow_pickle=False) as z:
        model=json.loads(str(z['model_json'])); names=[r['name'] for r in model['regions']]
        if args.region is None and len(names)!=1: p.error('--region is required for multiple regions.')
        index=0 if args.region is None else names.index(args.region); spec=model['regions'][index]; prefix=f'region_{index}_'; u=z[prefix+'u'];t=z[prefix+'t'];d=model['dimension']
        mesh=Mesh(z[prefix+'vertices'],z[prefix+'elements'],d,spec.get('element_order',model.get('element_order',2)),model.get('integration'))
    material=spec['material'];mat=Material(material['E'],material['nu'],d,model.get('state','plane_strain'),material.get('log_reference',mesh.scale)); points=np.atleast_2d(np.loadtxt(args.points,delimiter=',',skiprows=1)); policy=QuadraturePolicy(displacement_tol=args.u_tol,stress_tol=args.stress_tol)
    with threadpool_limits(limits=1): uv,sv,reports=fields(points,mesh,mat,u,t,policy)
    out=Path(args.out);out.mkdir(parents=True,exist_ok=True);np.savez_compressed(out/'fields.npz',points=points,u=uv,stress=sv)
    with open(out/'fields.csv','w',newline='',encoding='utf-8') as f:
        writer=csv.writer(f);writer.writerow(list('xyz'[:d])+['u'+a for a in 'xyz'[:d]]+['sxx','syy','szz','sxy','syz','szx','quadrature_converged','error_code','value_available'])
        for point,displacement,stress,r in zip(points,uv,sv,reports):writer.writerow([*point,*displacement,stress[0,0],stress[1,1],stress[2,2],stress[0,1],stress[1,2],stress[2,0],r['converged'],r['error_code'],r['value_available']])
    (out/'report.json').write_text(json.dumps(json_safe(reports),indent=2,allow_nan=False),encoding='utf-8');good=all(r['converged'] for r in reports)
    if not good: print('ERROR: accuracy unverified; available fields saved.')
    return 0 if good else 2
if __name__=='__main__':raise SystemExit(main())
