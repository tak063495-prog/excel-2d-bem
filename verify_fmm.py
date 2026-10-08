from pathlib import Path
import sys,json,time,copy,os
import numpy as np
import win32com.client
from threadpoolctl import threadpool_limits
ROOT=Path(__file__).resolve().parent;sys.stdout.reconfigure(encoding='utf-8');sys.path.insert(0,str(ROOT/'source/elastic_bem_unified'))
from elasticbem.model import solve_model
from elasticbem.material import Material
from elasticbem.meshes import Mesh
from elasticbem.operators import FMM,Dense
from elasticbem.tree import Chebyshev
from elasticbem.precision import fields,QuadraturePolicy
OUT=ROOT/'outputs/bem_vba_stage3';QA=ROOT/'qa_stage3';BOOK=OUT/'Elastic_BEM_VBA_Full.xlsm'
results=[]
def run(app,book,name,*args):return app.Run(f"'{book.Name}'!{name}",*args)
def good(app,book,name,*args):
    value=run(app,book,name,*args);assert not str(value).startswith('ERROR:'),(name,value);return value
def record(name,detail):results.append(dict(test=name,status='PASS',detail=detail));print('PASS',name,detail,flush=True)
def main():
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.AutomationSecurity=1;app.EnableEvents=False;book=None
    try:
        book=app.Workbooks.Open(str(BOOK),0,True);app.EnableEvents=True
        comp=book.VBProject.VBComponents.Add(1);comp.Name='QA_Fmm';comp.CodeModule.AddFromString((QA/'QA_Fmm.vba').read_text(encoding='utf-8'))
        for d in [2,3]:
            for p in [2,4,6]+([8] if d==2 else []):
                got=np.array(run(app,book,'QA_Cheb',p,d),float).ravel();ref=Chebyshev(p,d).basis(np.array([[.12,-.17,.61]])[:,:d],np.array([-.04,.02,.09])[:d],.8)[0]
                np.testing.assert_allclose(got,ref,rtol=2e-12,atol=1e-14);record(f'Chebyshev basis {d}D p={p}',{'max_difference':float(np.max(abs(got-ref)))})
        for n in [2,3,4,6]:
            for singular in [False,True]:
                i,j=np.indices((n,n));a=np.sin((i+1)*.73+(j+1)*.17)+2*np.eye(n)
                if singular:a[:,-1]=a[:,0]
                got=np.array(run(app,book,'QA_Pinv',n,singular),float);ref=np.linalg.pinv(a,rcond=1e-12)
                np.testing.assert_allclose(got,ref,rtol=2e-11,atol=2e-12);record(f'Jacobi SVD pseudoinverse n={n} singular={singular}',{'max_difference':float(np.max(abs(got-ref)))})
        cases=[
            ('2D_RECT','plane_strain',0,4,6,24,.7),('2D_RECT','plane_strain',2,16,6,24,.7),
            ('2D_RECT','plane_strain',2,64,6,24,.7),
            ('2D_RECT','plane_stress',2,12,6,8,.7),('2D_JOIN','plane_strain',2,8,6,8,.7),
            ('3D_BOX','three_dimensional',0,1,4,24,.7),('3D_BOX','three_dimensional',2,1,4,24,.7),
            ('3D_JOIN','three_dimensional',0,1,4,24,.7),('3D_JOIN','three_dimensional',2,1,4,24,.7),
            ('3D_BOX','three_dimensional',0,4,4,8,.9),
            ('3D_LONGBOX','three_dimensional',2,8,4,24,.7),
            ('2D_CIRCLE','plane_strain',2,24,6,8,.7),('3D_SPHERE','three_dimensional',0,8,4,8,.7)]
        if '--quick' in sys.argv:cases=cases[:2]
        if '--large-only' in sys.argv:cases=[case for case in cases if case[3]==64]
        if '--extra-3d' in sys.argv:cases=[case for case in cases if case[0]=='3D_LONGBOX']
        for kind,state,order,div,p,leaf,theta in cases:
            print('START',kind,state,order,div,flush=True);good(app,book,'API_Sample','3D_BOX' if kind=='3D_LONGBOX' else kind);op=book.Worksheets('操作');reg=book.Worksheets('領域');d=int(kind[0]);nreg=2 if 'JOIN' in kind else 1
            op.Range('B7').Value2=order;op.Range('B13').Value2='fmm';op.Range('B49').Value2=p;op.Range('B50').Value2=leaf;op.Range('B51').Value2=theta;op.Range('B58').Value2='gmres';op.Range('B43').Value2=1e-10
            if d==2:op.Range('B6').Value2=state
            for row in range(6,6+nreg):reg.Range(f'K{row}:M{row}').Value2=((div,div,div),)
            if 'SPHERE' in kind:reg.Range('L6').Value2=4
            if kind=='3D_LONGBOX':reg.Range('G6:I6').Value2=((8,1,1),);reg.Range('K6:M6').Value2=((8,1,1),)
            if 'CIRCLE' in kind or 'SPHERE' in kind:
                # Constant translation checks the non-zero prescribed displacement path on curved boundaries.
                book.Worksheets('境界条件').Range('D6').Value2=.001
                book.Worksheets('境界条件').Range('F6').Value2=-.002
                if d==3:book.Worksheets('境界条件').Range('H6').Value2=.003
            start=time.time();status=good(app,book,'API_Solve');elapsed=time.time()-start
            assert str(status).startswith('OK:'),(kind,status)
            model=json.loads(run(app,book,'ModelJSON'));name=f'{kind}_{state}_o{order}_n{div}_p{p}';(QA/(name+'.json')).write_text(json.dumps(model,indent=2),encoding='utf-8')
            stats=json.loads(run(app,book,'QA_State'));assert all(r['dense_arrays_empty'] for r in stats['regions']),stats
            with threadpool_limits(limits=1):regions,uu,tt,linear=solve_model(model)
            n=sum(len(r['mesh'].x) for r in regions);boundary=book.Worksheets('境界結果').Range(f'A6:M{5+n}').Value2;interior=book.Worksheets('内点結果').Range(f'A6:Q{5+nreg}').Value2
            offset=0;maxop=0.;maxu=0.;maxt=0.;far=0;maxs=0.
            for r,u,t in zip(regions,uu,tt):
                mesh=r['mesh'];nr=len(mesh.x);actual=np.array(boundary[offset:offset+nr],object);offset+=nr;far+=len(r['op'].far)
                ugot=np.array(actual[:,7:7+d],float);tgot=np.array(actual[:,10:10+d],float)
                np.testing.assert_allclose(ugot,u,rtol=4e-6,atol=1e-9);np.testing.assert_allclose(tgot,t,rtol=4e-6,atol=2e-5)
                maxu=max(maxu,float(np.max(abs(ugot-u))));maxt=max(maxt,float(np.max(abs(tgot-t))))
                for seed in [1.2,2.9]:
                    ids=np.arange(d*nr);pt=np.cos(ids*.29+seed).reshape(nr,d);pu=(.001*np.sin(ids*.37+seed)).reshape(nr,d)
                    got=np.array(run(app,book,'QA_Apply',r['name'],seed),float).ravel()
                    with threadpool_limits(limits=1):ref=r['op'].apply(pt,pu).ravel()
                    rel=float(np.max(abs(got-ref))/max(np.max(abs(ref)),1e-300));assert rel<2e-10,(name,rel)
                    maxop=max(maxop,rel)
                for component in range(1,d+1):
                    got=np.array(run(app,book,'QA_Apply',r['name'],0,component),float).ravel();assert np.max(abs(got))<2e-12,(name,got)
                for which,ref in [('G',r['op'].diagG),('H',r['op'].diagH)]:
                    got=np.array(run(app,book,'QA_Diagonal',r['name'],which),float)
                    np.testing.assert_allclose(got,ref,rtol=3e-9,atol=1e-12)
                # Far pairs must be real tree interactions, independently of the selector label.
                got=run(app,book,'QA_Plan',r['name'],'far');got=np.array(got,int).reshape(-1,2) if got is not None else np.empty((0,2),int)
                ref=np.array([(a.number,b.number) for a,b in r['op'].far],int).reshape(-1,2)
                assert np.array_equal(got,ref),(name,'far plan differs',got.shape,ref.shape)
                row=next(row for row in interior if row[0]==r['name'])
                with threadpool_limits(limits=1):uv,sv,checks=fields(r['spec']['interior_points'],mesh,r['mat'],u,t,QuadraturePolicy(**model['quadrature']))
                selected=np.array([sv[0,0,0],sv[0,1,1],sv[0,2,2],sv[0,0,1],sv[0,1,2],sv[0,2,0]])
                np.testing.assert_allclose(np.array(row[8:14],float),selected,rtol=4e-6,atol=3e-5);maxs=max(maxs,float(np.max(abs(np.array(row[8:14],float)-selected))))
            if kind=='2D_RECT' and div>=12:assert far>0,(name,far)
            if kind=='3D_BOX' and div==4:assert far>0,(name,far)
            if kind=='3D_LONGBOX':assert far>0 and n*d>1200,(name,far,n*d)
            record(name,dict(DOF=n*d,far_pairs=far,seconds=elapsed,VBA_iterations=stats['iterations'],VBA_residual=stats['residual'],Python_residual=linear['true_relative_residual'],operator_relative_max_difference=maxop,boundary_u_max_difference_m=maxu,boundary_t_max_difference_kPa=maxt,interior_stress_max_difference_kPa=maxs,dense_arrays_empty=True))
        if '--quick' not in sys.argv and '--large-only' not in sys.argv and '--extra-3d' not in sys.argv:
            report=dict(test_count=len(results),all_passed=True,reference='Original elasticbem.operators.FMM and elasticbem.model.solve_model',tests=results)
            (OUT/'FMM検証結果.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        if '--large-only' in sys.argv:(OUT/'FMM大規模検証結果.json').write_text(json.dumps(dict(test_count=len(results),all_passed=True,tests=results),ensure_ascii=False,indent=2),encoding='utf-8')
        if '--extra-3d' in sys.argv:(OUT/'FMM3D二次要素検証結果.json').write_text(json.dumps(dict(test_count=len(results),all_passed=True,tests=results),ensure_ascii=False,indent=2),encoding='utf-8')
        print('ALL PASSED',len(results),flush=True)
    finally:
        app.EnableEvents=False
        if book is not None:book.Close(False)
        app.Quit()
if __name__=='__main__':main()
