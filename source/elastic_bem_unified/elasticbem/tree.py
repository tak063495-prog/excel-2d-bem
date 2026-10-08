"""Shared dimension-independent quadtree/octree and Chebyshev interpolation."""
from dataclasses import dataclass,field
import numpy as np
@dataclass
class Node:
    ids: np.ndarray
    center: np.ndarray
    half0: float
    half: float
    level: int
    number: int
    children: list=field(default_factory=list)
class Chebyshev:
    def __init__(self,p,d):
        if not isinstance(p,int) or not 2<=p<=(8 if d==2 else 6): raise ValueError('FMM p: 2..8 (2D), 2..6 (3D).')
        self.p=p; self.dimension=d; self.z=np.cos(np.pi*(np.arange(p)+.5)/p); self.bw=(-1.)**np.arange(p)*np.sin(np.pi*(np.arange(p)+.5)/p); self.grid=np.stack(np.meshgrid(*([self.z]*d),indexing='ij'),axis=-1).reshape(-1,d); self.q=p**d
    def basis(self,x,center,half):
        xx=(x-center)/half; result=np.ones((len(x),1))
        for k in range(self.dimension):
            delta=xx[:,k,None]-self.z; close=abs(delta)<1e-13; a=self.bw/np.where(close,1,delta); a/=a.sum(axis=1,keepdims=True); rows=close.any(axis=1); a[rows]=close[rows].astype(float); result=(result[:,:,None]*a[:,None,:]).reshape(len(x),-1)
        return result
class InteractionPlan:
    def __init__(self,mesh,leaf=24,theta=.7):
        if leaf<1 or not 0<theta<1: raise ValueError('Invalid leaf/theta.')
        self.mesh=mesh; self.nodes=[]; self.leaves=[]; self.far=[]; self.near=[]; d=mesh.dimension
        lo=mesh.v.min(axis=(0,1)); hi=mesh.v.max(axis=(0,1)); half0=max(hi-lo)/2*(1+1e-12); pad=np.max(abs(mesh.v-mesh.x[:,None]))*(1+1e-12)
        bits=2**np.arange(d)
        def build(ids,c,h,level):
            node=Node(ids,c,h,h+pad,level,len(self.nodes)); self.nodes.append(node)
            if len(ids)>leaf and level<20:
                code=((mesh.x[ids]>c).astype(int)*bits).sum(axis=1)
                for octant in np.unique(code):
                    selected=ids[code==octant]; side=((octant&bits)>0).astype(int); node.children.append(build(selected,c+(2*side-1)*h/2,h/2,level+1))
            else: self.leaves.append(node)
            return node
        self.root=build(np.arange(len(mesh.x)),(lo+hi)/2,half0,0)
        def walk(a,b):
            distance=np.linalg.norm(a.center-b.center)
            if distance>0 and np.sqrt(d)*(a.half+b.half)<theta*distance: self.far.append((a,b))
            elif not a.children and not b.children: self.near.append((a,b))
            elif a.children and (not b.children or a.half0>=b.half0):
                for c in a.children: walk(c,b)
            else:
                for c in b.children: walk(a,c)
        walk(self.root,self.root); self.near_entries=sum(len(a.ids)*len(b.ids) for a,b in self.near); self.translation_count=len({self.key(a,b) for a,b in self.far})
    def key(self,a,b): return (a.level,b.level,tuple(np.round((b.center-a.center)/self.root.half0,13)))
