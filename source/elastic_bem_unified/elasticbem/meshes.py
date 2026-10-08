"""Finite closed line2/line3 and tri3/tri6 boundary geometry.
Field interpolation is constant or discontinuous quadratic.
"""
import numpy as np
from scipy.optimize import minimize
from .rules import gauss, gauss01, de01, triangle_gauss, triangle_self, triangle_polar, closest_triangle

LINE_NODES=np.array([[-2/3],[0.],[2/3]])
TRI_GEOM=np.array([[0.,0.],[1.,0.],[0.,1.],[.5,0.],[.5,.5],[0.,.5]])
TRI_FIELD=np.array([1/3,1/3])+.65*(TRI_GEOM-np.array([1/3,1/3]))
def monomial(s):
    a,b=np.asarray(s).T; return np.c_[np.ones(len(a)),a,b,a*a,a*b,b*b]
TRI_INV=np.linalg.inv(monomial(TRI_FIELD))
def geom_line(s,order):
    a=np.asarray(s)[:,0]
    if order==0: return np.c_[(1-a)/2,(1+a)/2],np.tile([-.5,.5],(len(a),1))
    return np.c_[a*(a-1)/2,1-a*a,a*(a+1)/2],np.c_[a-.5,-2*a,a+.5]
def geom_tri(s,order):
    a,b=np.asarray(s).T; c=1-a-b
    if order==0:
        return np.c_[c,a,b],np.tile(np.array([[-1,1,0],[-1,0,1]])[None],(len(a),1,1))
    N=np.c_[c*(2*c-1),a*(2*a-1),b*(2*b-1),4*c*a,4*a*b,4*b*c]
    ds=np.c_[1-4*c,4*a-1,np.zeros(len(a)),4*(c-a),4*b,-4*b]
    dt=np.c_[1-4*c,np.zeros(len(a)),4*b-1,-4*a,4*a,4*(c-b)]
    return N,np.stack([ds,dt],axis=1)

class Mesh:
    def __init__(self,vertices,elements,dimension=2,element_order=2,integration=None):
        self.dimension=d=dimension; self.order=element_order; self.vertices=np.asarray(vertices,float); self.elements=np.asarray(elements)
        if d not in (2,3) or element_order not in (0,2): raise ValueError('dimension 2/3, element_order 0/2 required.')
        if self.vertices.ndim!=2 or self.vertices.shape[1]!=d or not np.isfinite(self.vertices).all(): raise ValueError('Invalid finite vertices array.')
        required=(3 if element_order==2 else 2) if d==2 else (6 if element_order==2 else 3)
        if self.elements.ndim!=2 or self.elements.shape[1]!=required or len(self.elements)<(3 if d==2 else 4) or not np.issubdtype(self.elements.dtype,np.integer): raise ValueError('Invalid element connectivity shape/type.')
        if self.elements.min()<0 or self.elements.max()>=len(self.vertices): raise ValueError('Invalid geometry index.')
        self.curves=self.vertices[self.elements]; self.ne=len(self.elements); self.nfield=1 if self.order==0 else (3 if d==2 else 6)
        self.param_nodes=(np.array([[0.]]) if d==2 else np.array([[1/3,1/3]])) if self.order==0 else (LINE_NODES if d==2 else TRI_FIELD)
        self.scale=float(np.ptp(self.vertices,axis=0).max())
        if self.scale<=0: raise ValueError('Zero mesh scale.')
        integration=integration or {}; self.boundary_order=int(integration.get('boundary_order',48 if d==2 else 12)); self.boundary_de_level=int(integration.get('boundary_de_level',3 if d==2 else 1)); self.self_order=int(integration.get('self_order',24))
        if not 4<=self.boundary_order<=128 or not 0<=self.boundary_de_level<=4 or not 8<=self.self_order<=128: raise ValueError('Invalid boundary integration controls.')
        self.element_length=[]; support=[]; self.planar=[]
        for e,c in enumerate(self.curves):
            if d==2:
                endpoint=c[-1]; length=np.linalg.norm(endpoint-c[0]); probes=np.linspace(-1,1,33)[:,None]
                if self.order==2:
                    a=(c[0]+c[2])/2-c[1]; b=(c[2]-c[0])/2; ss=np.clip(-a@b/max(2*a@a,1e-300),-1,1)
                    if np.linalg.norm(2*a*ss+b)<1e-12*self.scale: raise ValueError('Degenerate line Jacobian.')
                    control=np.array([c[0],2*c[1]-(c[0]+c[2])/2,c[2]])
                else: control=c
                planar=self.order==0 or np.max(abs(c[1]-(c[0]+c[2])/2))<1e-12*self.scale
            else:
                length=max(np.linalg.norm(c[0]-c[1]),np.linalg.norm(c[0]-c[2]),np.linalg.norm(c[1]-c[2])); probes=triangle_gauss(12)[0]
                if self.order==2: control=np.r_[c[:3],np.array([2*c[3]-(c[0]+c[1])/2,2*c[4]-(c[1]+c[2])/2,2*c[5]-(c[2]+c[0])/2])]
                else: control=c
                planar=self.order==0 or np.max(abs(c[3:]-np.array([(c[0]+c[1])/2,(c[1]+c[2])/2,(c[2]+c[0])/2])))<1e-12*self.scale
            y,j,n=self.geometry(e,probes)
            if length<1e-12*self.scale or j.min()<1e-14*self.scale**(d-1): raise ValueError('Degenerate geometry.')
            if d==3:
                base=np.cross(c[1]-c[0],c[2]-c[0]); base/=np.linalg.norm(base)
                if np.any(n@base<=0): raise ValueError('Folded or overturned quadratic triangle.')
            support.append([control.min(axis=0),control.max(axis=0)]); self.element_length.append(length); self.planar.append(planar)
        self.element_length=np.array(self.element_length); self.v=np.repeat(np.array(support),self.nfield,axis=0)
        self.x=np.concatenate([self.geometry(e,self.param_nodes)[0] for e in range(self.ne)]); self.normal=np.concatenate([self.geometry(e,self.param_nodes)[2] for e in range(self.ne)])
        self._validate_closed()
    def shape(self,s):
        if self.order==0: return np.ones((len(s),1))
        if self.dimension==3: return monomial(s)@TRI_INV
        a=np.asarray(s)[:,0]; nodes=LINE_NODES[:,0]
        return np.stack([np.prod([(a-nodes[j])/(nodes[i]-nodes[j]) for j in range(3) if j!=i],axis=0) for i in range(3)],axis=1)
    def geometry(self,e,s):
        if self.dimension==2:
            N,dN=geom_line(s,self.order); y=N@self.curves[e]; derivative=dN@self.curves[e]; j=np.linalg.norm(derivative,axis=1); n=np.c_[derivative[:,1],-derivative[:,0]]/np.maximum(j[:,None],1e-300)
        else:
            N,dN=geom_tri(s,self.order); y=N@self.curves[e]; derivative=np.einsum('qak,kd->qad',dN,self.curves[e]); cross=np.cross(derivative[:,0],derivative[:,1]); j=np.linalg.norm(cross,axis=1); n=cross/np.maximum(j[:,None],1e-300)
        return y,j,n
    def derivatives(self,e,s):
        if self.dimension==2: return (geom_line(s,self.order)[1]@self.curves[e])[:,None,:]
        return np.einsum('qak,kd->qad',geom_tri(s,self.order)[1],self.curves[e])
    def regular_rule(self,order=None):
        if self.dimension==2:
            z,w=gauss(order or self.boundary_order); return z[:,None],w
        return triangle_gauss(order or self.boundary_order)
    def _validate_closed(self):
        if self.dimension==2:
            ends=self.elements[:,[0,-1]]; mapping={int(a):int(b) for a,b in ends}
            if len(mapping)!=self.ne or set(ends[:,0])!=set(ends[:,1]): raise ValueError('Require one closed oriented loop.')
            p=int(ends[0,0]); seen=[]
            for _ in ends: seen.append(p); p=mapping[p]
            if p!=seen[0] or len(set(seen))!=self.ne: raise ValueError('Require one connected loop.')
            area=0.
            for e in range(self.ne):
                s,w=self.regular_rule(24); y,j,n=self.geometry(e,s); area+=w@(np.sum(y*n,axis=1)*j)
            if area<=0: raise ValueError('2D finite boundary must be counterclockwise.')
        else:
            edges={}; adjacent=[set() for _ in self.elements]
            for k,t in enumerate(self.elements):
                for a,b,mid in [(t[0],t[1],t[3] if self.order==2 else -1),(t[1],t[2],t[4] if self.order==2 else -1),(t[2],t[0],t[5] if self.order==2 else -1)]: edges.setdefault(tuple(sorted([int(a),int(b)])),[]).append((k,1 if a<b else -1,int(mid)))
            for edge,uses in edges.items():
                if len(uses)!=2 or uses[0][1]+uses[1][1]!=0 or uses[0][2]!=uses[1][2]: raise ValueError('3D mesh must be watertight; share midside indices and orient outward.')
                a,b=uses[0][0],uses[1][0]; adjacent[a].add(b); adjacent[b].add(a)
            seen={0}; todo=[0]
            while todo:
                for k in adjacent[todo.pop()]:
                    if k not in seen: seen.add(k); todo.append(k)
            if len(seen)!=self.ne: raise ValueError('One connected surface required per region.')
            volume=0.
            for e in range(self.ne):
                s,w=triangle_gauss(10); y,j,n=self.geometry(e,s); volume+=w@(np.sum(y*n,axis=1)*j)/3
            if volume<=0: raise ValueError('Orient 3D surface outward (positive signed volume).')
    def closest(self,e,x):
        c=self.curves[e]
        if self.dimension==2:
            if self.order==0:
                v=c[1]-c[0]; a=np.clip((x-c[0])@v/(v@v),0,1); s=2*a-1; return np.array([s]),float(np.linalg.norm(c[0]+a*v-x))
            a=(c[0]+c[2])/2-c[1]; b=(c[2]-c[0])/2; z=c[1]-x; coeff=[2*a@a,3*a@b,b@b+2*a@z,b@z]; roots=np.roots(np.trim_zeros(coeff,'f')) if np.any(coeff) else []
            values=[-1.,1.]+[float(r.real) for r in roots if abs(r.imag)<1e-10 and -1<r.real<1]; ss=np.array(values)[:,None]; dist=np.linalg.norm(self.geometry(e,ss)[0]-x,axis=1); i=np.argmin(dist); return ss[i],float(dist[i])
        z,st,d=closest_triangle(x,c[:3])
        if self.planar[e]: return st,d
        def fun(s):
            r=self.geometry(e,s[None])[0][0]-x; der=self.derivatives(e,s[None])[0]; return r@r,2*der@r
        seeds=[st, np.array([1/3,1/3])]; best=None
        for seed in seeds:
            result=minimize(fun,seed,jac=True,method='SLSQP',bounds=[(0,1),(0,1)],constraints=[dict(type='ineq',fun=lambda a:1-a.sum(),jac=lambda a:-np.ones(2))],options=dict(ftol=1e-20,maxiter=100))
            p=np.clip(result.x,0,1); p/=max(1,p.sum()); val=fun(p)[0]
            if best is None or val<best[0]: best=(val,p)
        return best[1],float(np.sqrt(best[0]))
    def integration_rule(self,e,x,method='auto',level=None,self_param=None,order=None):
        level=self.boundary_de_level if level is None else level
        if method=='gauss': return self.regular_rule(order)
        if self_param is not None:
            if self.dimension==3: return triangle_self(self_param,self.self_order)
            s=float(self_param[0]); g,w=de01(level); points=[]; weights=[]
            for sign,extent in [(-1,s+1),(1,1-s)]:
                if extent<=0: continue
                a=s+sign*extent*g; keep=abs(a-s)>1e-12; points.append(a[keep,None]); weights.append(extent*w[keep])
            return np.concatenate(points),np.concatenate(weights)
        st,d=self.closest(e,x)
        if method=='auto' and d/self.element_length[e]>.2: return self.regular_rule(order)
        speed=np.linalg.norm(self.derivatives(e,st[None])[0],ord=2); delta=d/max(speed,1e-300)
        if self.dimension==3: return triangle_polar(st,delta,level,method='de' if method in ('auto','de') else 'asinh')
        g,w=de01(level) if method in ('auto','de') else gauss01(24*2**min(level,5)); z=[]; weights=[]; s=st[0]
        for sign,extent in [(-1,s+1),(1,1-s)]:
            if extent<=0: continue
            upper=np.arcsinh(extent/max(delta,1e-15)); a=delta*np.sinh(upper*g); z.append((s+sign*a)[:,None]); weights.append(delta*np.cosh(upper*g)*upper*w)
        return np.concatenate(z),np.concatenate(weights)
    def inside(self,x):
        if self.dimension==2:
            p=np.concatenate([self.geometry(e,np.linspace(-1,1,65)[:-1,None])[0] for e in range(self.ne)])-x; q=np.roll(p,-1,axis=0); winding=np.sum(np.arctan2(p[:,0]*q[:,1]-p[:,1]*q[:,0],np.sum(p*q,axis=1)))
            if min(self.closest(e,x)[1]/self.element_length[e] for e in range(self.ne))>.01: return abs(winding)>np.pi
        else:
            # Signed solid angle of geometry corners suffices for planar surfaces.
            if all(self.planar):
                r=self.curves[:,:3]-x; a,b,c=r[:,0],r[:,1],r[:,2]; la=np.linalg.norm(a,axis=1); lb=np.linalg.norm(b,axis=1); lc=np.linalg.norm(c,axis=1)
                angle=2*np.arctan2(np.einsum('ij,ij->i',a,np.cross(b,c)),la*lb*lc+np.sum(a*b,axis=1)*lc+np.sum(b*c,axis=1)*la+np.sum(c*a,axis=1)*lb)
                return abs(angle.sum())>2*np.pi
        # Curved surface winding via mapped Gauss/DE flux, rather than chord containment.
        flux=0.
        for e in range(self.ne):
            st,d=self.closest(e,x); method='de' if d/self.element_length[e]<.2 else 'gauss'; s,w=self.integration_rule(e,x,method=method,level=2,order=24 if self.dimension==2 else 12); y,j,n=self.geometry(e,s); r=y-x; norm=np.linalg.norm(r,axis=1); flux+=w@(j*np.sum(r*n,axis=1)/norm**self.dimension)
        return abs(flux)>(np.pi if self.dimension==2 else 2*np.pi)
