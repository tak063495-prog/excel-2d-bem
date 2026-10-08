"""Convert original constant NPZ meshes/BCs or finite 2D JSON to unified input."""
import argparse,json
from pathlib import Path
import numpy as np

def main():
    p=argparse.ArgumentParser(description=__doc__); p.add_argument('mesh'); p.add_argument('bc',nargs='?'); p.add_argument('--out',default='converted.json'); p.add_argument('--E',type=float,default=30000); p.add_argument('--nu',type=float,default=.3); p.add_argument('--state',choices=['plane_strain','plane_stress'],default='plane_strain'); args=p.parse_args()
    if Path(args.mesh).suffix.lower()=='.json':
        model=json.loads(Path(args.mesh).read_text(encoding='utf-8-sig'))
        if any(r.get('kind','finite')!='finite' for r in model['regions']): raise ValueError('Infinite-domain models cannot be converted into finite domains automatically.')
        model['dimension']=len(model['regions'][0]['vertices'][0]); model.setdefault('element_order',2)
        q=model.get('quadrature',{})
        for key in ['allow_unconverged','rtol','atol']: q.pop(key,None)
    else:
        if args.bc is None: p.error('BC NPZ is required for mesh NPZ.')
        with np.load(args.mesh,allow_pickle=False) as z:
            vertices=z['vertices']; d=vertices.shape[1]; elements=z['segments' if d==2 else 'triangles']
        with np.load(args.bc,allow_pickle=False) as z: mask=z['is_displacement']; values=z['values']
        if mask.dtype!=bool or mask.shape!=(len(elements),d) or values.shape!=mask.shape or not np.isfinite(values).all(): raise ValueError('Invalid constant-element BC arrays.')
        bcs=[dict(elements=[i],type=['u' if m else 't' for m in row],values=values[i].tolist()) for i,row in enumerate(mask)]
        model=dict(dimension=d,element_order=0,regions=[dict(name='body',vertices=vertices.tolist(),elements=elements.tolist(),material=dict(E=args.E,nu=args.nu),boundary_conditions=bcs)],solver=dict(backend='auto'))
        if d==2: model['state']=args.state
    Path(args.out).write_text(json.dumps(model,indent=2),encoding='utf-8'); print(args.out)
if __name__=='__main__': main()
