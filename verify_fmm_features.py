from pathlib import Path
import sys,json,time,zipfile
import numpy as np
import win32com.client
from threadpoolctl import threadpool_limits
from verify_fmm import ROOT,OUT,QA,BOOK,run,good
from elasticbem.meshes import Mesh
from elasticbem.material import Material
from elasticbem.operators import FMM
from elasticbem.io import vtk
tests=[]
def record(name,detail):tests.append(dict(test=name,status='PASS',detail=detail));print('PASS',name,detail,flush=True)
def configure(app,book,backend='fmm',order=2,div=12):
    good(app,book,'API_Sample','2D_RECT');op=book.Worksheets('操作')
    op.Range('B13').Value2=backend;op.Range('B7').Value2=order;op.Range('B58').Value2='gmres';op.Range('B43').Value2=1e-10
    op.Range('B49').Value2=4;op.Range('B50').Value2=8;op.Range('B51').Value2=.7;op.Range('B52').Value2=128;op.Range('B53').Value2=512
    op.Range('B54').Value2=120;op.Range('B55').Value2=100;op.Range('B56').Value2=200;op.Range('B57').Value2=30000
    op.Range('B59').Value2='auto';op.Range('B60').Value2=.2;op.Range('B61').Value2=1e-10
    book.Worksheets('領域').Range('K6:M6').Value2=((div,div,1),)
    return op
def compare_npz(folder,book):
    m=json.loads((folder/'model.json').read_text(encoding='utf-8'));d=m['dimension'];spec=m['regions'][0];mesh=Mesh(spec['vertices'],spec['elements'],d,m['element_order'],m['integration'])
    boundary=np.array(book.Worksheets('境界結果').Range(f'A6:M{5+len(mesh.x)}').Value2,object);row=book.Worksheets('内点結果').Range('A6:Q6').Value2[0]
    with zipfile.ZipFile(folder/'solution.npz') as z:assert z.testzip() is None
    with np.load(folder/'solution.npz',allow_pickle=False) as data:
        assert json.loads(str(data['model_json']))==m and int(data['dimension'])==d
        np.testing.assert_allclose(data['region_0_vertices'],mesh.vertices,atol=1e-13);np.testing.assert_array_equal(data['region_0_elements'],mesh.elements)
        np.testing.assert_allclose(data['region_0_collocation'],mesh.x,atol=1e-13)
        np.testing.assert_allclose(data['region_0_u'],np.array(boundary[:,7:7+d],float),atol=1e-16)
        np.testing.assert_allclose(data['region_0_t'],np.array(boundary[:,10:10+d],float),atol=1e-13)
        if row[16]:
            np.testing.assert_allclose(data['region_0_interior_u'][0],np.array(row[5:5+d],float),atol=1e-16)
            np.testing.assert_allclose(data['region_0_interior_stress'][0],[[row[8],row[11],row[13]],[row[11],row[9],row[12]],[row[13],row[12],row[10]]],atol=1e-12)
        else:assert np.isnan(data['region_0_interior_u']).all() and np.isnan(data['region_0_interior_stress']).all()
    expected=folder/'expected.vtk';vtk(expected,mesh,np.array(boundary[:,7:7+d],float),np.array(boundary[:,10:10+d],float))
    lines=(folder/'block_boundary.vtk').read_text(encoding='utf-8').splitlines();ref=expected.read_text(encoding='ascii').splitlines()
    assert len(lines)==len(ref)
    for a,b in zip(lines,ref):
        try:aa=[float(x) for x in a.split()];bb=[float(x) for x in b.split()]
        except ValueError:assert a==b,(a,b);continue
        np.testing.assert_allclose(aa,bb,rtol=3e-12,atol=1e-12)
    return m
def main():
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.AutomationSecurity=1;app.EnableEvents=False;book=None
    try:
        book=app.Workbooks.Open(str(BOOK),0,True);app.EnableEvents=True
        comp=book.VBProject.VBComponents.Add(1);comp.Name='QA_Fmm';comp.CodeModule.AddFromString((QA/'QA_Fmm.vba').read_text(encoding='utf-8'))
        for backend,method,cap in [('dense','gmres',1200),('dense','lu',1200),('auto','gmres',1200),('auto','gmres',12)]:
            op=configure(app,book,backend,2,4);op.Range('B58').Value2=method;op.Range('B34').Value2=cap
            status=good(app,book,'API_Solve');assert str(status).startswith('OK:'),status
            stats=json.loads(run(app,book,'QA_State'));actual=stats['regions'][0]['operator']['backend']
            if cap==12:assert actual=='fmm' and stats['regions'][0]['dense_arrays_empty']
            if backend=='dense':assert actual=='dense' and not stats['regions'][0]['dense_arrays_empty']
            np.testing.assert_allclose(book.Worksheets('内点結果').Range('I6:K6').Value2[0],[100,0,30],atol=3e-5)
            record(f'{backend} / {method} dense cap={cap}',{'selected':actual,'residual':stats['residual'],'iterations':stats['iterations']})
        op.Range('B34').Value2=1200
        for cache in [0,.035]:
            op=configure(app,book);op.Range('B52').Value2=cache;good(app,book,'API_Solve');m=json.loads(run(app,book,'ModelJSON'));spec=m['regions'][0];mesh=Mesh(spec['vertices'],spec['elements'],2,2,m['integration']);mat=Material(30000,.3,2,m['state'],mesh.scale)
            with threadpool_limits(limits=1):ref=FMM(mesh,mat,p=4,leaf=8,theta=.7,cache_mb=cache)
            for seed in [1.2,2.9]:
                ids=np.arange(2*len(mesh.x));t=np.cos(ids*.29+seed).reshape(-1,2);u=(.001*np.sin(ids*.37+seed)).reshape(-1,2)
                got=np.array(run(app,book,'QA_Apply','block',seed),float).reshape(-1,2)
                with threadpool_limits(limits=1):expected=ref.apply(t,u)
                np.testing.assert_allclose(got,expected,rtol=2e-9,atol=1e-12)
            stats=json.loads(run(app,book,'QA_State'))['regions'][0]['operator'];assert stats['m2l_cache_MB']<=cache+1e-9 and stats['far_pairs']>0
            record(f'FMM bounded LRU cache={cache} MB',stats)
        for d,order in [(2,0),(2,2),(3,0),(3,2)]:
            op=configure(app,book,'fmm',order,4)
            if d==3:
                good(app,book,'API_Sample','3D_BOX');op.Range('B7').Value2=order;book.Worksheets('領域').Range('K6:M6').Value2=((1,1,1),)
            op.Range('B52').Value2=128;good(app,book,'API_Solve');folder=QA/f'full_export_{d}d_o{order}';folder.mkdir(exist_ok=True);good(app,book,'API_Export',str(folder));compare_npz(folder,book)
            report=json.loads((folder/'report.json').read_text(encoding='utf-8'));assert report['all_requested_checks_passed'] and report['regions'][0]['interior_reports'][0]['element_reports']
            # CSV, binary arrays and VTK all refer to the same solution, not an earlier Python run.
            good(app,book,'API_Import',str(folder));record(f'VBA CSV/NPZ/VTK export {d}D order={order}',{'report_method':report['linear_solver']['method'],'schema':'original Python compatible'})
        op=configure(app,book);op.Range('B55').Value2=1;op.Range('B56').Value2=1
        status=good(app,book,'API_Solve');assert str(status).startswith('UNMET:'),status
        folder=QA/'gmres_unmet';folder.mkdir(exist_ok=True);good(app,book,'API_Export',str(folder));report=json.loads((folder/'report.json').read_text(encoding='utf-8'))
        assert report['linear_solver']['gmres_info']>0 and report['linear_solver']['error_code']=='LINEAR_SOLVE_TOLERANCE_UNMET' and not report['all_requested_checks_passed']
        compare_npz(folder,book);record('GMRES iteration limit retains computed solution',report['linear_solver'])
        op=configure(app,book);op.Range('B42').Value2=1;status=good(app,book,'API_Solve');assert str(status).startswith('UNMET:')
        folder=QA/'quadrature_unavailable';folder.mkdir(exist_ok=True);good(app,book,'API_Export',str(folder));compare_npz(folder,book)
        report=json.loads((folder/'report.json').read_text(encoding='utf-8'));assert not report['regions'][0]['interior_reports'][0]['value_available']
        good(app,book,'API_Import',str(folder));record('unavailable internal fields preserved in NPZ/CSV/report','NaN arrays and False flags')
        op.Range('B42').Value2=200000
        for mode in ['gauss','de']:
            op=configure(app,book,'dense',2,4);op.Range('B59').Value2=mode;status=good(app,book,'API_Solve');assert str(status).startswith('OK:')
            np.testing.assert_allclose(book.Worksheets('内点結果').Range('I6:K6').Value2[0],[100,0,30],atol=.001);record('explicit internal quadrature '+mode,status)
        op=configure(app,book)
        for cell,bad in [('B49',1),('B50',0),('B51',1),('B52',-1),('B53',0),('B54',0),('B55',0),('B56',0),('B57',1),('B58','lu'),('B59','bad'),('B60',0),('B61',0)]:
            previous=op.Range(cell).Value2;op.Range(cell).Value2=bad;rejected=str(run(app,book,'API_Solve'));assert rejected.startswith('ERROR:'),(cell,rejected);record('invalid FMM/GMRES setting '+cell,rejected);op.Range(cell).Value2=previous
        op.Range('B53').Value2=1;rejected=str(run(app,book,'API_Solve'));assert rejected.startswith('ERROR:') and 'メモリ' in rejected;record('FMM model memory rejection before near allocation',rejected)
        op.Range('B53').Value2=512
        good(app,book,'API_Sample','3D_BOX');op.Range('B49').Value2=7;rejected=str(run(app,book,'API_Solve'));assert rejected.startswith('ERROR:');record('3D p upper bound',rejected)
        report=dict(test_count=len(tests),all_passed=True,tests=tests);(OUT/'演算選択・出力検証結果.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print('ALL PASSED',len(tests),flush=True)
    finally:
        app.EnableEvents=False
        if book is not None:book.Close(False)
        app.Quit()
if __name__=='__main__':main()
