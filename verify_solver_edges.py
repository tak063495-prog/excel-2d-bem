from pathlib import Path
import sys,json,time,csv
import numpy as np
import win32com.client
from threadpoolctl import threadpool_limits
from verify_solver import ROOT,OUT,QA,BOOK,run,good
from elasticbem.meshes import Mesh
from elasticbem.material import Material
from elasticbem.precision import fields,QuadraturePolicy
tests=[]
def record(name,detail):tests.append(dict(test=name,status='PASS',detail=detail));print('PASS',name,detail,flush=True)
def main():
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.AutomationSecurity=1;app.EnableEvents=False;book=None
    try:
        book=app.Workbooks.Open(str(BOOK),0,True);app.EnableEvents=True
        comp=book.VBProject.VBComponents.Add(1);comp.Name='QA_Solver';comp.CodeModule.AddFromString((QA/'QA_Solver.vba').read_text(encoding='utf-8'))
        for kind,order in [('2D_RECT',2),('2D_CIRCLE',2),('3D_BOX',0),('3D_SPHERE',2)]:
            good(app,book,'API_Sample',kind);d=int(kind[0]);op=book.Worksheets('操作');op.Range('B7').Value2=order;reg=book.Worksheets('領域')
            if kind=='3D_BOX':reg.Range('K6:M6').Value2=((1,1,1),)
            if kind=='3D_SPHERE':reg.Range('K6:M6').Value2=((8,4,1),)
            if kind=='2D_CIRCLE':reg.Range('K6').Value2=8
            good(app,book,'API_Generate',False);m=json.loads(run(app,book,'ModelJSON'));spec=m['regions'][0];mesh=Mesh(spec['vertices'],spec['elements'],d,order,m['integration'])
            if kind=='2D_RECT':point=np.array([.9999,.43])
            elif kind=='3D_BOX':point=np.array([.999,.43,.37])
            else:
                p=mesh.geometry(0,np.array([[.27]]) if d==2 else np.array([[.23,.31]]))[0][0];point=p*.998
            book.Worksheets('評価点').Range('C6:E6').Value2=(tuple([*point,*([0.] if d==2 else [])]),)
            good(app,book,'API_Generate',False);m=json.loads(run(app,book,'ModelJSON'));mat=Material(30000,.3,d,m['state'],mesh.scale)
            print('START near',kind,flush=True)
            actual=np.array(run(app,book,'QA_Closest'),float)
            with threadpool_limits(limits=1):closest=[mesh.closest(e,point) for e in range(mesh.ne)]
            diff=max(abs(actual[e,2]-distance) for e,(st,distance) in enumerate(closest));assert diff<2e-8,(kind,diff)
            record(f'closest point near {kind}',{'max_distance_difference_m':diff})
            start=time.time();rows=run(app,book,'QA_InitAffine');row=rows[0]
            ex=100/mat.E;ey=-mat.nu*ex
            if mat.dimension==2 and mat.state=='plane_strain':ex*=1-mat.nu**2;ey=-mat.nu*(1+mat.nu)*100/mat.E
            strain=np.diag([ex,ey]+([ey] if d==3 else []));stress=mat.stress(strain);u=mesh.x@strain.T;t=mesh.normal@stress[:d,:d].T
            with threadpool_limits(limits=1):uv,sv,checks=fields([point],mesh,mat,u,t,QuadraturePolicy(**m['quadrature']))
            assert row[16] and row[14]==checks[0]['converged'],(kind,row,checks)
            np.testing.assert_allclose(np.array(row[5:5+d],float),uv[0],rtol=3e-6,atol=1e-9)
            selected=[sv[0,0,0],sv[0,1,1],sv[0,2,2],sv[0,0,1],sv[0,1,2],sv[0,2,0]]
            np.testing.assert_allclose(np.array(row[8:14],float),selected,rtol=3e-6,atol=.002)
            record(f'near-boundary adaptive fields {kind}',{'stress_max_difference_kPa':float(np.max(abs(np.array(row[8:14],float)-selected))),'converged':bool(row[14]),'seconds':time.time()-start})
        # Native solver CSVs use the Python schema and round-trip through the existing importer.
        for d in [2,3]:
            good(app,book,'API_Sample','2D_RECT' if d==2 else '3D_BOX');op.Range('B7').Value2=0
            if d==3:reg.Range('K6:M6').Value2=((1,1,1),)
            good(app,book,'API_Solve');folder=QA/f'vba_export_{d}d';folder.mkdir(exist_ok=True)
            n=int(op.Range('B21').Value2/d);before=np.array(book.Worksheets('境界結果').Range(f'A6:M{5+n}').Value2,object)
            good(app,book,'API_Export',str(folder));report=json.loads((folder/'report.json').read_text(encoding='utf-8-sig'))
            assert report['all_requested_checks_passed'] and report['linear_solver']['converged'],report
            data=np.loadtxt(folder/'block_boundary.csv',delimiter=',',skiprows=1,encoding='utf-8-sig')
            np.testing.assert_allclose(data[:,:d],np.array(before[:,4:4+d],float),atol=1e-13)
            np.testing.assert_allclose(data[:,d:2*d],np.array(before[:,7:7+d],float),rtol=1e-13,atol=1e-17)
            good(app,book,'API_Import',str(folder));after=np.array(book.Worksheets('境界結果').Range(f'A6:M{5+n}').Value2,object)
            for start in [4,7,10]:np.testing.assert_allclose(np.array(after[:,start:start+d],float),np.array(before[:,start:start+d],float),rtol=1e-13,atol=1e-13)
            record(f'VBA solve/export/import {d}D',{'relative_residual':report['linear_solver']['true_relative_residual'],'CSV_schema':'Python compatible'})
        good(app,book,'API_Sample','2D_RECT');op.Range('B7').Value2=2;reg.Range('G6:H6').Value2=((.001,.001),)
        book.Worksheets('評価点').Range('C6:D6').Value2=((.0005,.0005),);book.Worksheets('材料').Range('B6').Value2=3e7
        status=good(app,book,'API_Solve');row=book.Worksheets('内点結果').Range('A6:Q6').Value2[0]
        np.testing.assert_allclose(row[8:14],[100,0,30,0,0,0],atol=.001)
        record('geometry/material scaling',{'size_m':.001,'E_kPa':3e7,'sxx_kPa':row[8],'status':status})
        report=dict(test_count=len(tests),all_passed=True,tests=tests)
        (OUT/'境界近傍・入出力検証結果.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print('ALL PASSED',len(tests),flush=True)
    finally:
        app.EnableEvents=False
        if book is not None:book.Close(False)
        app.Quit()
if __name__=='__main__':main()
