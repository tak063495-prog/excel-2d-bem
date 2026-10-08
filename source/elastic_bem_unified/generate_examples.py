"""Editable finite-domain ND examples. No infinite-domain examples."""
import json
from pathlib import Path
import numpy as np

def box_mesh(d,n=1,lower=None,upper=None,order=2):
    lower=np.zeros(d) if lower is None else np.array(lower,float); upper=np.ones(d) if upper is None else np.array(upper,float)
    if n<1 or np.any(upper<=lower): raise ValueError('Positive box sizes and n>=1 required.')
    vertices=[]; elements=[]; faces=[[] for _ in range(2*d)]; lookup={}
    def vertex(ijk):
        key=tuple(ijk)
        if key not in lookup: lookup[key]=len(vertices); vertices.append(lower+(upper-lower)*np.array(ijk)/n)
        return lookup[key]
    if d==2:
        corners=np.array([[0,0],[n,0],[n,n],[0,n]]); face_order=[2,1,3,0]
        for side in range(4):
            for i in range(n):
                a=corners[side]+(corners[(side+1)%4]-corners[side])*i/n; b=corners[side]+(corners[(side+1)%4]-corners[side])*(i+1)/n; ia=vertex(a); ib=vertex(b)
                if order==2: im=vertex((a+b)/2); element=[ia,im,ib]
                else: element=[ia,ib]
                faces[face_order[side]].append(len(elements)); elements.append(element)
    else:
        for axis in range(3):
            others=[j for j in range(3) if j!=axis]
            for side in [0,n]:
                normal=np.zeros(3); normal[axis]=-1 if side==0 else 1
                for i in range(n):
                    for j in range(n):
                        pts=[]
                        for a,b in [(i,j),(i+1,j),(i+1,j+1),(i,j+1)]:
                            p=np.zeros(3); p[axis]=side; p[others]=[a,b]; pts.append(p)
                        for indices in [(0,1,2),(0,2,3)]:
                            p=np.array([pts[k] for k in indices])
                            if np.cross(p[1]-p[0],p[2]-p[0])@normal<0: p=p[[0,2,1]]
                            ids=[vertex(a) for a in p]
                            if order==2: ids += [vertex((p[0]+p[1])/2),vertex((p[1]+p[2])/2),vertex((p[2]+p[0])/2)]
                            faces[2*axis+(side==n)].append(len(elements)); elements.append(ids)
    return np.array(vertices),np.array(elements),faces

def region(name,d,n=1,E=30000,nu=.3,lower=None,upper=None,order=2):
    v,e,f=box_mesh(d,n,lower,upper,order); return dict(name=name,vertices=v.tolist(),elements=e.tolist(),material=dict(E=E,nu=nu),element_order=order),f

def block_model(d=3,state='plane_strain',n=1,order=2):
    r,faces=region('block',d,n,order=order); bcs=[]
    for j in range(d):
        kinds=['t']*d; kinds[j]='u'; bcs.append(dict(elements=faces[2*j],type=kinds,values=[0.]*d))
    load=[100.]+[0.]*(d-1); bcs.append(dict(elements=faces[1],type=['t']*d,values=load)); r['boundary_conditions']=bcs; r['interior_points']=[[.5]*d]
    model=dict(dimension=d,element_order=order,regions=[r],solver=dict(backend='auto'))
    if d==2: model['state']=state; r['interior_points'].append([.9999,.4])
    return model

def two_material_model(d=3,state='plane_strain',n=1,order=2):
    nub=.3 if d==3 or state=='plane_stress' else (-1+np.sqrt(1+8*.15*1.15))/2
    a,fa=region('left',d,n,nu=.15,order=order); lo=np.zeros(d); hi=np.ones(d); lo[0]=1.; hi[0]=2.; b,fb=region('right',d,n,E=60000,nu=nub,lower=lo,upper=hi,order=order)
    a['boundary_conditions']=[]; b['boundary_conditions']=[]
    for j in range(d):
        kinds=['t']*d; kinds[j]='u'; a['boundary_conditions'].append(dict(elements=fa[2*j],type=kinds,values=[0.]*d))
        if j>0: b['boundary_conditions'].append(dict(elements=fb[2*j],type=kinds,values=[0.]*d))
    b['boundary_conditions'].append(dict(elements=fb[1],type=['t']*d,values=[100.]+[0.]*(d-1)))
    # Pair interface elements by their unordered physical corner sets.
    va=np.array(a['vertices']); vb=np.array(b['vertices']); ea=np.array(a['elements']); eb=np.array(b['elements']); matched=[]
    for e in fa[1]:
        ca=va[ea[e]][[0,-1]] if d==2 else va[ea[e],:][:3]
        for f in fb[0]:
            cb=vb[eb[f]][[0,-1]] if d==2 else vb[eb[f],:][:3]
            if np.max(np.min(np.linalg.norm(ca[:,None]-cb[None],axis=-1),axis=1))<1e-10: matched.append(f); break
        else: raise ValueError('Generated nonconforming triangulation.')
    a['interior_points']=[[.5]*d]; b['interior_points']=[[1.5]+[.5]*(d-1)]
    model=dict(dimension=d,element_order=order,regions=[a,b],interfaces=[dict(region_a='left',elements_a=fa[1],region_b='right',elements_b=matched)],solver=dict(backend='auto'))
    if d==2: model['state']=state
    return model

def main():
    out=Path(__file__).parent/'examples'; out.mkdir(exist_ok=True)
    for d,state in [(2,'plane_strain'),(2,'plane_stress'),(3,'three_dimensional')]:
        for name,fn in [('block',block_model),('two_materials',two_material_model)]:
            for order in [0,2]:
                model=fn(d,state,n=2 if d==2 else 1,order=order); (out/f'{name}_{d}d_order{order}_{state}.json').write_text(json.dumps(model,indent=2),encoding='utf-8')
if __name__=='__main__': main()
