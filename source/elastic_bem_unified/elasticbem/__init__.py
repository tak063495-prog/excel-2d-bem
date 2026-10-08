"""Unified 2D/3D finite-domain elasticity BEM."""
from .material import Material
from .meshes import Mesh
from .operators import Dense,FMM,choose_operator
from .model import solve_model
from .precision import QuadraturePolicy,fields
__all__=['Material','Mesh','Dense','FMM','choose_operator','solve_model','QuadraturePolicy','fields']
