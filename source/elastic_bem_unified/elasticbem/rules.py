"""Shared Gauss / tanh-sinh DE and natural-triangle polar rules."""
from functools import lru_cache
import numpy as np
from scipy.special import expit
from numpy.polynomial.legendre import leggauss

@lru_cache(maxsize=32)
def gauss(order): return leggauss(order)
@lru_cache(maxsize=32)
def gauss01(order):
    z,w=gauss(order); return (z+1)/2,w/2
@lru_cache(maxsize=12)
def de01(level):
    step=.5/2**level; t=np.arange(-int(4/step),int(4/step)+1)*step; z=np.pi*np.sinh(t); a=expit(z); b=expit(-z); w=step*np.pi*np.cosh(t)*a*b
    keep=(a>0)&(a<1)&(w>0); return a[keep],w[keep]
@lru_cache(maxsize=32)
def triangle_gauss(order):
    a,w=gauss01(order); s,t=np.meshgrid(a,a,indexing='ij'); ws,wt=np.meshgrid(w,w,indexing='ij')
    return np.c_[s.ravel(),((1-s)*t).ravel()],(ws*wt*(1-s)).ravel()

def closest_triangle(x,v):
    # Interior plane projection and all three edge projections; no active-set ambiguity.
    A=np.stack([v[1]-v[0],v[2]-v[0]],axis=1); st=np.linalg.lstsq(A,x-v[0],rcond=None)[0]
    candidates=[]
    if min(st)>=0 and sum(st)<=1: candidates.append(v[0]+A@st)
    for a,b in zip(v,np.roll(v,-1,axis=0)):
        t=np.clip((x-a)@(b-a)/((b-a)@(b-a)),0,1); candidates.append(a+t*(b-a))
    z=min(candidates,key=lambda y:np.linalg.norm(y-x)); st=np.linalg.lstsq(A,z-v[0],rcond=None)[0]
    return z,st,float(np.linalg.norm(z-x))

def triangle_self(st,order=24):
    a,wa=gauss01(order); b,wb=gauss01(max(48,order)); z=[]; w=[]
    for x,y in zip(np.array([[0.,0.],[1.,0.],[0.,1.]]),np.array([[1.,0.],[0.,1.],[0.,0.]])):
        p=x-st; q=y-st; det=abs(p[0]*q[1]-p[1]*q[0]); direction=(1-b[:,None])*p+b[:,None]*q
        z.append((st+a[:,None,None]*direction[None]).reshape(-1,2)); w.append((a[:,None]*wa[:,None]*wb[None]*det).ravel())
    return np.concatenate(z),np.concatenate(w)

def triangle_polar(st,distance,level,method='de'):
    """Natural triangle: nearest-point fan and angular asinh mapping.
    Physical curvature enters geometry mapping/Jacobian after this rule.
    """
    v=np.array([[0.,0.],[1.,0.],[0.,1.]]); z=np.asarray(st); d=max(distance,1e-15)
    radial,rw=de01(level) if method=='de' else gauss01(24*2**min(level,3))
    ag,aw=gauss01(12*2**min(level,4) if method=='de' else 24*2**min(level,3))
    points=[]; weights=[]
    for a,b in zip(v,np.roll(v,-1,axis=0)):
        ra=a-z; rb=b-z; edge=b-a; outward=np.array([edge[1],-edge[0]])/np.linalg.norm(edge); height=ra@outward
        if height<1e-14: continue
        tangent=np.array([-outward[1],outward[0]]); wa=np.arcsinh(ra@tangent/height); wb=np.arcsinh(rb@tangent/height)
        if wb<=wa: raise ValueError('Invalid polar partition.')
        angle=wa+(wb-wa)*ag; cp=1/np.cosh(angle); sp=np.tanh(angle); direction=cp[:,None]*outward+sp[:,None]*tangent; limit=height*np.cosh(angle)
        angular=aw*(wb-wa)*cp
        if method=='de':
            A=np.log1p((limit/d)**2); av=A[:,None]*radial[None]; rho=d*np.sqrt(np.expm1(av)); radial_jac=.5*d*d*A[:,None]*np.exp(av)*rw[None]
        else:
            upper=np.arcsinh(limit/d); s=upper[:,None]*radial[None]; rho=d*np.sinh(s); radial_jac=d*d*np.sinh(s)*np.cosh(s)*upper[:,None]*rw[None]
        points.append((z+rho[:,:,None]*direction[:,None]).reshape(-1,2)); weights.append((angular[:,None]*radial_jac).ravel())
    if not points: raise ValueError('Polar partition failed.')
    return np.concatenate(points),np.concatenate(weights)
