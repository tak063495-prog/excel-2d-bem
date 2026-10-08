from pathlib import Path
import sys,json,time,copy,warnings
import numpy as np
import win32com.client
from threadpoolctl import threadpool_limits
ROOT=Path(__file__).resolve().parent
sys.stdout.reconfigure(encoding='utf-8');sys.path.insert(0,str(ROOT/'source/elastic_bem_unified'))
from elasticbem.model import solve_model
from elasticbem.material import Material,kernels,gradients
from elasticbem.precision import fields,QuadraturePolicy
STAGE3='--stage3' in sys.argv
OUT=ROOT/('outputs/bem_vba_stage3' if STAGE3 else 'outputs/bem_vba_stage2');QA=ROOT/('qa_stage3/near_checks' if STAGE3 else 'qa_stage2');BOOK=OUT/('Elastic_BEM_VBA_Full.xlsm' if STAGE3 else 'Elastic_BEM_VBA_Solver.xlsm')
if STAGE3:
    import shutil
    QA.mkdir(parents=True,exist_ok=True);shutil.copy2(ROOT/'qa_stage2/QA_Solver.vba',QA/'QA_Solver.vba')
results=[]
def record(name,detail):
    results.append(dict(test=name,status='PASS',detail=detail));print('PASS',name,detail,flush=True)
def run(app,book,name,*args):return app.Run(f"'{book.Name}'!{name}",*args)
def good(app,book,name,*args):
    r=run(app,book,name,*args);assert not str(r).startswith('ERROR:'),(name,r);return r
def main():
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.AutomationSecurity=1;app.EnableEvents=False
    book=None
    try:
        book=app.Workbooks.Open(str(BOOK),0,True);app.EnableEvents=True
        comp=book.VBProject.VBComponents.Add(1);comp.Name='QA_Solver';comp.CodeModule.AddFromString((QA/'QA_Solver.vba').read_text(encoding='utf-8'))
        # Quadrature and analytic target gradients, independently of boundary solution.
        for n in [4,12,24,48,112]:
            a=np.array(run(app,book,'QA_Gauss',n),float)
            err=max(abs(np.sum(a[:,1]*a[:,0]**i)-1/(i+1)) for i in range(min(2*n,30)))
            assert err<2e-14,(n,err);record(f'Gauss-Legendre n={n}',{'max_moment_error':err})
        for d,state in [(2,'plane_strain'),(2,'plane_stress'),(3,'three_dimensional')]:
            good(app,book,'API_Sample','2D_RECT' if d==2 else '3D_BOX')
            if d==2:book.Worksheets('操作').Range('B6').Value2=state
            good(app,book,'API_Generate',False)
            actual=np.array(str(run(app,book,'QA_Kernel')).split(','),float)
            mat=Material(30000,.3,d,state,.8);x=np.array([.12,-.17,.09])[:d];y=np.array([[.78,.28,.61]])[:,:d];n=np.array([[.6,.8]]) if d==2 else np.array([[1/3,2/3,2/3]])
            u,ds=kernels(x[None],y,mat);t=np.einsum('qijk,qk->qij',ds[0],n);du,dt=gradients(x,y,n,mat)
            ref=np.r_[u.ravel(),t.ravel(),du.ravel(),dt.ravel()];np.testing.assert_allclose(actual,ref,rtol=2e-13,atol=2e-17)
            # Central differences substantiate the analytic gradients in the supplied Python too.
            for axis in range(d):
                step=np.zeros(d);step[axis]=1e-6
                plus,dp=kernels((x+step)[None],y,mat);minus,dm=kernels((x-step)[None],y,mat)
                np.testing.assert_allclose((plus-minus)[0]/2e-6,du[:,:,:,axis],rtol=2e-8,atol=1e-13)
                np.testing.assert_allclose(np.einsum('qijk,qk->qij',(dp-dm)[0],n)/2e-6,dt[:,:,:,axis],rtol=2e-8,atol=1e-10)
            record(f'Kelvin U,T,dU,dT {d}D {state}',{'max_abs_difference':float(np.max(abs(actual-ref)))})
        cases=[]
        for d in [2,3]:
            states=['plane_strain','plane_stress'] if d==2 else ['three_dimensional']
            for state in states:
                for order in [0,2]:cases.append(('2D_RECT' if d==2 else '3D_BOX',state,order))
            for order in [0,2]:cases.append(('2D_JOIN' if d==2 else '3D_JOIN','plane_strain' if d==2 else 'three_dimensional',order))
        for shape in ['2D_CIRCLE','2D_POLYGON','3D_SPHERE']:
            for order in [0,2]:cases.append((shape,'plane_strain' if shape.startswith('2D') else 'three_dimensional',order))
        for kind,state,order in cases:
            good(app,book,'API_Sample',kind);op=book.Worksheets('操作');reg=book.Worksheets('領域')
            d=int(kind[0]);op.Range('B7').Value2=order
            if d==2:op.Range('B6').Value2=state
            count=2 if 'JOIN' in kind else 1
            if 'CIRCLE' in kind:reg.Range('K6').Value2=8
            elif 'SPHERE' in kind:reg.Range('K6:M6').Value2=((8,4,1),)
            elif 'POLYGON' not in kind:
                for row in range(6,6+count):reg.Range(f'K{row}:M{row}').Value2=((2,2,1),) if d==2 else ((1,1,1),)
            # Non-zero rigid translation is exact on any curved surface and checks RHS decoding.
            if 'CIRCLE' in kind or 'SPHERE' in kind or 'POLYGON' in kind:
                book.Worksheets('境界条件').Range('D6').Value2=.001
                book.Worksheets('境界条件').Range('F6').Value2=-.002
                if d==3:book.Worksheets('境界条件').Range('H6').Value2=.003
            good(app,book,'API_Generate',False)
            name=f'{kind}_{state}_order{order}';print('START',name,flush=True);start=time.time()
            status=good(app,book,'API_Solve');elapsed=time.time()-start
            assert str(status).startswith('OK:'),(name,status)
            model=json.loads(run(app,book,'ModelJSON'));(QA/(name+'.json')).write_text(json.dumps(model,indent=2),encoding='utf-8')
            reference=copy.deepcopy(model);reference['solver']['backend']='dense';reference['linear_solver']['rtol']=1e-11
            with threadpool_limits(limits=1):regions,uu,tt,report=solve_model(reference)
            if not any(s in kind for s in ['CIRCLE','SPHERE','POLYGON']):assert report['converged'],report
            n=sum(len(r['mesh'].x) for r in regions)
            boundary=book.Worksheets('境界結果').Range(f'A6:M{5+n}').Value2
            interior=book.Worksheets('内点結果').Range(f'A6:Q{5+count}').Value2
            offset=0;stats=dict(seconds=elapsed,DOF=n*d,python_residual=report['true_relative_residual']);gerr=0.;herr=0.;uerr=0.;terr=0.;ferr=0.
            for r,u,t in zip(regions,uu,tt):
                mesh=r['mesh'];nr=len(mesh.x);actual=np.array(boundary[offset:offset+nr],object);offset+=nr
                np.testing.assert_allclose(np.array(actual[:,4:4+d],float),mesh.x,atol=1e-13)
                ugot=np.array(actual[:,7:7+d],float);tgot=np.array(actual[:,10:10+d],float)
                np.testing.assert_allclose(ugot,u,rtol=2e-7,atol=5e-10)
                np.testing.assert_allclose(tgot,t,rtol=2e-7,atol=5e-6)
                uerr=max(uerr,float(np.max(abs(ugot-u))));terr=max(terr,float(np.max(abs(tgot-t))))
                for which,ref in [('G',r['op'].Gmatrix),('H',r['op'].Hmatrix)]:
                    got=np.array(run(app,book,'QA_Matrix',r['name'],which),float)
                    rel=float(np.max(abs(got-ref))/max(np.max(abs(ref)),1e-300))
                    assert rel<3e-8,(name,which,rel)
                    if which=='G':gerr=max(gerr,rel)
                    else:herr=max(herr,rel)
                points=r['spec']['interior_points'];pol=QuadraturePolicy(**model['quadrature'])
                with threadpool_limits(limits=1):uv,sv,checks=fields(points,mesh,r['mat'],u,t,pol)
                row=next(row for row in interior if row[0]==r['name'])
                np.testing.assert_allclose(np.array(row[5:5+d],float),uv[0],rtol=3e-7,atol=5e-10)
                selected=np.array([sv[0,0,0],sv[0,1,1],sv[0,2,2],sv[0,0,1],sv[0,1,2],sv[0,2,0]])
                np.testing.assert_allclose(np.array(row[8:14],float),selected,rtol=3e-7,atol=3e-5)
                assert row[14]==checks[0]['converged'] and row[16]==checks[0]['value_available'],(name,row,checks)
                ferr=max(ferr,float(np.max(abs(np.array(row[8:14],float)-selected))))
                if order==2 and kind in ['2D_RECT','3D_BOX']:
                    E=r['mat'].E;nu=r['mat'].nu;ex=100/E;ey=-nu*ex
                    if state=='plane_strain':ex*=(1-nu**2);ey=-nu*(1+nu)*100/E
                    expected=np.array([ex*.5,ey*.5]+([ey*.5] if d==3 else []))
                    np.testing.assert_allclose(np.array(row[5:5+d],float),expected,atol=1e-8)
                    np.testing.assert_allclose(np.array(row[8:14],float),[100,0,30 if state=='plane_strain' else 0,0,0,0],atol=1e-3)
            stats.update(G_relative_max_difference=gerr,H_relative_max_difference=herr,boundary_u_max_difference_m=uerr,boundary_t_max_difference_kPa=terr,interior_stress_max_difference_kPa=ferr,VBA_status=status)
            record(name,stats)
            # Postprocessor with exact affine boundary values on curved quadratic geometry.
            if order==2 and ('CIRCLE' in kind or 'SPHERE' in kind):
                affine=run(app,book,'QA_AffineFields');arow=affine[0]
                assert arow[14] and arow[16],arow
                # On quadratic curved geometry strain is polynomial, while traction includes a non-polynomial normal.
                # Compare against Python using exactly the same discontinuous nodal boundary fields.
                r=regions[0];strain=np.zeros((d,d));ex=100/r['mat'].E;ey=-r['mat'].nu*ex
                if state=='plane_strain':ex*=(1-r['mat'].nu**2);ey=-r['mat'].nu*(1+r['mat'].nu)*100/r['mat'].E
                np.fill_diagonal(strain,[ex,ey]+([ey] if d==3 else []));sig=r['mat'].stress(strain)
                ua=r['mesh'].x@strain.T;ta=r['mesh'].normal@sig[:d,:d].T
                with threadpool_limits(limits=1):av,asig,ac=fields(r['spec']['interior_points'],r['mesh'],r['mat'],ua,ta,pol)
                selected=[asig[0,0,0],asig[0,1,1],asig[0,2,2],asig[0,0,1],asig[0,1,2],asig[0,2,0]]
                np.testing.assert_allclose(np.array(arow[8:14],float),selected,rtol=3e-7,atol=3e-5)
                record(name+' affine internal fields',{'stress_max_difference_kPa':float(np.max(abs(np.array(arow[8:14],float)-selected)))})
        # Bounds, atomic stale status, and explicit precision-failure behavior.
        good(app,book,'API_Sample','2D_RECT');good(app,book,'API_Generate',False)
        op.Range('B34').Value2=12
        rejected=str(run(app,book,'API_Solve'));assert rejected.startswith('ERROR:') and '上限' in rejected
        record('dense DOF cap before allocation',rejected);op.Range('B34').Value2=1200
        for cell,value in [('B35',2),('B38',0),('B40',1),('B42',0),('B43',1)]:
            original=op.Range(cell).Value2;op.Range(cell).Value2=value
            rejected=str(run(app,book,'API_Solve'));assert rejected.startswith('ERROR:'),(cell,rejected)
            record('invalid solve setting '+cell,rejected);op.Range(cell).Value2=original
        op.Range('B38').Value2=1e-30;op.Range('B39').Value2=1e-30;op.Range('B42').Value2=32
        status=good(app,book,'API_Solve');row=book.Worksheets('内点結果').Range('A6:Q6').Value2[0]
        assert str(status).startswith('UNMET:') and row[14] is False and row[16] is True and isinstance(row[8],float)
        record('precision unmet retains finite values',{'status':status,'error_code':row[15],'value_available':row[16]})
        op.Range('B42').Value2=1;status=good(app,book,'API_Solve');row=book.Worksheets('内点結果').Range('A6:Q6').Value2[0]
        assert str(status).startswith('UNMET:') and row[16] is False and row[8]=='取得不可'
        record('budget exhausted preserves unavailable flag',{'status':status,'error_code':row[15]})
        op.Range('B38').Value2=1e-8;op.Range('B39').Value2=.001;op.Range('B42').Value2=200000
        good(app,book,'API_Solve');book.Worksheets('材料').Range('B6').Value2=40000
        assert '旧入力' in op.Range('B22').Value2
        record('changed input invalidates solver results',op.Range('B22').Value2)
        # Keep user deliverable free of temporary QA components and altered test models.
        report=dict(test_count=len(results),all_passed=True,tests=results,reference='Attached elastic_bem_unified Python dense implementation',scope='Numerical implementation agreement; not mesh convergence or physical-model certification')
        (OUT/'解析核検証結果.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print('ALL PASSED',len(results),flush=True)
    finally:
        app.EnableEvents=False
        if book is not None:book.Close(False)
        app.Quit()
if __name__=='__main__':main()
