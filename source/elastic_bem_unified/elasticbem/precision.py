"""Shared physical-tolerance Gauss -> DE interior evaluation, retains failures."""
from dataclasses import dataclass
import time,warnings
import numpy as np
from .material import kernels,gradients
@dataclass(frozen=True)
class QuadraturePolicy:
    mode: str='auto'
    displacement_tol: float=1e-8
    stress_tol: float=1e-3
    near_ratio: float=.2
    max_de_level: int=6
    max_gauss_level: int=6
    max_points: int=200000
    min_distance_ratio: float=1e-10
    def __post_init__(self):
        if self.mode not in ('auto','gauss','de') or not np.isfinite([self.displacement_tol,self.stress_tol,self.near_ratio,self.min_distance_ratio]).all() or min(self.displacement_tol,self.stress_tol,self.near_ratio,self.min_distance_ratio)<=0 or self.max_points<1 or min(self.max_de_level,self.max_gauss_level)<2: raise ValueError('Invalid quadrature policy.')

def json_safe(v):
    if isinstance(v,dict): return {k:json_safe(a) for k,a in v.items()}
    if isinstance(v,(list,tuple)): return [json_safe(a) for a in v]
    if isinstance(v,np.ndarray): return json_safe(v.tolist())
    if isinstance(v,(float,np.floating)): return float(v) if np.isfinite(v) else None
    if isinstance(v,np.integer): return int(v)
    if isinstance(v,np.bool_): return bool(v)
    return v

def fields(points,mesh,mat,u,t,policy=None):
    policy=policy or QuadraturePolicy(); points=np.asarray(points,float); d=mesh.dimension; k=mesh.nfield
    if points.ndim!=2 or points.shape[1]!=d or not np.isfinite(points).all(): raise ValueError('Invalid finite interior point array.')
    if np.asarray(u).shape!=mesh.x.shape or np.asarray(t).shape!=mesh.x.shape or not np.isfinite(u).all() or not np.isfinite(t).all(): raise ValueError('Invalid finite boundary fields.')
    outu=[]; outs=[]; reports=[]; ns=d*d+(1 if d==2 else 0); target=np.r_[np.full(d,policy.displacement_tol),np.full(ns,policy.stress_tol)]
    for x in points:
        start=time.perf_counter(); geoms=[mesh.closest(e,x) for e in range(mesh.ne)]
        valid=min(distance/mesh.element_length[e] for e,(st,distance) in enumerate(geoms))>policy.min_distance_ratio and mesh.inside(x)
        if not valid:
            outu.append(np.full(d,np.nan)); outs.append(np.full((3,3),np.nan)); reports.append(dict(converged=False,error_code='INVALID_INTERIOR_POINT',value_available=False,point=x.tolist())); continue
        nearest=int(np.argmin([distance for st,distance in geoms])); uref=mesh.shape(geoms[nearest][0][None])[0]@u[k*nearest:k*(nearest+1)]; total=np.zeros(d+ns); error=np.zeros(d+ns); details=[]; unavailable=False
        for e,(st,distance) in enumerate(geoms):
            eta=distance/mesh.element_length[e]; method='de' if policy.mode=='de' or (policy.mode=='auto' and eta<policy.near_ratio) else 'gauss'; count=0; value=None; estimate=None; good=False; switched=False; reason='refinement_limit_or_roundoff'
            def evaluate(method,level):
                nonlocal count
                order=([8,12,20,32,48,72,112][min(level,6)] if d==2 else [4,6,8,12,16,24,32][min(level,6)])*2**max(level-6,0)
                upper=order**(d-1) if method=='gauss' else ((16*2**level+1)*2 if d==2 else 3*(16*2**level+1)*12*2**min(level,4))
                if count+upper>policy.max_points: return None,None
                s,w=mesh.integration_rule(e,x,method=method,level=level,order=order)
                if count+len(w)>policy.max_points: return None,None
                y,j,n=mesh.geometry(e,s); N=mesh.shape(s); uy=N@u[k*e:k*(e+1)]; ty=N@t[k*e:k*(e+1)]; U,D=kernels(x[None],y,mat); T=np.einsum('qijk,qk->qij',D[0],n); uv=np.einsum('qij,qj->qi',U[0],ty)-np.einsum('qij,qj->qi',T,uy-uref)
                dU,dT=gradients(x,y,n,mat); grad=np.einsum('qijl,qj->qil',dU,ty)-np.einsum('qijl,qj->qil',dT,uy-uref); stress=mat.stress((grad+grad.transpose(0,2,1))/2)
                f=np.c_[uv,stress[:,:d,:d].reshape(-1,d*d)]
                if d==2: f=np.c_[f,stress[:,2,2]]
                weights=w*j; count+=len(w); return weights@f,64*np.finfo(float).eps*(abs(weights)@abs(f))
            for family in ([method,'de'] if method=='gauss' and policy.mode=='auto' else [method]):
                if family!=method: switched=True
                method=family; previous=None; passed=0
                for level in range((policy.max_gauss_level if method=='gauss' else policy.max_de_level)+1):
                    v,floor=evaluate(method,level)
                    if v is None: reason='point_budget'; break
                    value=v
                    if previous is not None:
                        estimate=np.maximum(abs(v-previous),floor); passed=passed+1 if np.all(estimate<=target/mesh.ne) else 0
                        if passed>=2: good=True; reason='two_refinements'; break
                    previous=v.copy()
                if good: break
            if value is None: unavailable=True
            else: total+=value
            error+=estimate if estimate is not None else np.full(d+ns,np.inf)
            details.append(dict(element=e,method=method,converged=good,reason=reason,evaluations=count,gauss_to_de=switched,distance_ratio=eta,error_estimate=estimate.tolist() if estimate is not None else None))
        good=not unavailable and all(a['converged'] for a in details) and np.all(error<=target); uv=total[:d]+uref; stress=np.zeros((3,3)); stress[:d,:d]=total[d:d+d*d].reshape(d,d)
        if d==2: stress[2,2]=total[-1]
        if unavailable: uv[:]=np.nan; stress[:]=np.nan
        report=dict(converged=bool(good),error_code='OK' if good else 'QUADRATURE_TOLERANCE_UNMET',value_available=not unavailable,point=x.tolist(),error_estimate=error.tolist(),tolerance=target.tolist(),element_reports=details,de_elements=sum(a['method']=='de' for a in details),seconds=time.perf_counter()-start,error_scope='quadrature estimate only; boundary discretization/FMM/algebraic errors excluded'); reports.append(report); outu.append(uv); outs.append(stress)
        if not good: warnings.warn('QUADRATURE_TOLERANCE_UNMET: computed values retained.',RuntimeWarning)
    return np.array(outu),np.array(outs),reports
