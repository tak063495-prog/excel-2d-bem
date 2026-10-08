"""Shared material and ND Kelvin kernels. Units: kN, m, kPa."""
from dataclasses import dataclass
import numpy as np

@dataclass(frozen=True)
class Material:
    E: float=30000.
    nu: float=.3
    dimension: int=3
    state: str='plane_strain'
    log_reference: float=1.
    def __post_init__(self):
        if self.dimension not in (2,3) or not np.isfinite([self.E,self.nu,self.log_reference]).all() or self.E<=0 or not -1<self.nu<.5 or self.log_reference<=0: raise ValueError('dimension 2/3; E>0, -1<nu<.5, log_reference>0 required.')
        if self.dimension==2 and self.state not in ('plane_strain','plane_stress'): raise ValueError('2D state plane_strain/plane_stress required.')
    @property
    def mu(self): return self.E/(2*(1+self.nu))
    @property
    def effective_nu(self): return self.nu/(1+self.nu) if self.dimension==2 and self.state=='plane_stress' else self.nu
    @property
    def lam(self):
        return 2*self.mu*self.effective_nu/(1-2*self.effective_nu)
    def stress(self,strain):
        e=np.asarray(strain); trace=np.trace(e,axis1=-2,axis2=-1); out=np.zeros(e.shape[:-2]+(3,3))
        out[...,:self.dimension,:self.dimension]=2*self.mu*e+self.lam*trace[...,None,None]*np.eye(self.dimension)
        if self.dimension==2 and self.state=='plane_strain': out[...,2,2]=self.lam*trace
        return out

def kernels(x,y,mat):
    r=np.asarray(y)[None]-np.asarray(x)[:,None]; length=np.linalg.norm(r,axis=-1)
    if np.any(length<=0): raise ValueError('Coincident kernel coordinates: use singular integration.')
    d=mat.dimension; h=r/length[...,None]; I=np.eye(d); nu=mat.effective_nu; hh=h[..., :,None]*h[...,None,:]
    if d==2: U=(-(3-4*nu)*I*np.log(length[...,None,None]/mat.log_reference)+hh)/(8*np.pi*mat.mu*(1-nu))
    else: U=((3-4*nu)*I+hh)/(16*np.pi*mat.mu*(1-nu)*length[...,None,None])
    D=(1-2*nu)*(np.einsum('ij,abk->abijk',I,h)-np.einsum('abi,jk->abijk',h,I)+np.einsum('abj,ik->abijk',h,I))+d*np.einsum('abi,abj,abk->abijk',h,h,h)
    D*=-1/((8 if d==3 else 4)*np.pi*(1-nu)*length[...,None,None,None]**(d-1))
    return U,D

def gradients(x,y,n,mat):
    d=mat.dimension; r=y-x; length=np.linalg.norm(r,axis=1); h=r/length[:,None]; I=np.eye(d); nu=mat.effective_nu; alpha=1-2*nu
    hh=h[:,:,None]*h[:,None,:]; hn=np.sum(h*n,axis=1)
    F=alpha*(I*hn[:,None,None]-h[:,:,None]*n[:,None,:]+h[:,None,:]*n[:,:,None])+d*hh*hn[:,None,None]
    dU=np.empty((len(y),d,d,d)); dT=np.empty_like(dU)
    for l in range(d):
        dh=I[:,l][None]-h*h[:,l,None]; dhn=n[:,l]-hn*h[:,l]
        a=(3-4*nu)*I*h[:,l,None,None]-np.einsum('i,qj->qij',I[:,l],h)-np.einsum('qi,j->qij',h,I[:,l])+d*hh*h[:,l,None,None]
        dU[:,:,:,l]=a/((16 if d==3 else 8)*np.pi*mat.mu*(1-nu)*length[:,None,None]**(d-1))
        dF=alpha*(I*dhn[:,None,None]-dh[:,:,None]*n[:,None,:]+dh[:,None,:]*n[:,:,None])
        dF+=d*((dh[:,:,None]*h[:,None,:]+h[:,:,None]*dh[:,None,:])*hn[:,None,None]+hh*dhn[:,None,None])
        dT[:,:,:,l]=(dF-(d-1)*F*h[:,l,None,None])/((8 if d==3 else 4)*np.pi*(1-nu)*length[:,None,None]**d)
    return dU,dT
