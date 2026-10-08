"""Additional native Excel checks for state, mixed operators and edge inputs."""
from pathlib import Path
import sys, json, copy, time
from unittest.mock import patch
import numpy as np
import win32com.client
from threadpoolctl import threadpool_limits
from verify_fmm import ROOT, OUT, QA, BOOK, run, good
from elasticbem.model import solve_model
from elasticbem.operators import Dense, FMM

tests = []
EXTRA_QA = '''
Public Function QA_Invalidated() As Boolean
QA_Invalidated = Len(LastSolvedFingerprint) = 0 And (SolveData Is Nothing) And (FieldReports Is Nothing)
End Function
Public Function QA_BusyState() As Boolean
QA_BusyState = Busy
End Function
'''

def record(name, detail):
    tests.append(dict(test=name, status='PASS', detail=detail))
    print('PASS', name, detail, flush=True)

def reset(app, book, kind='2D_RECT', div=4):
    good(app, book, 'API_Sample', kind)
    op = book.Worksheets('操作')
    op.Range('B13').Value2 = 'fmm'; op.Range('B34').Value2 = 1200
    op.Range('B35:B43').Value2 = ((0,), (24,), (-1,), (1e-8,), (.001,), (6,), (6,), (200000,), (1e-10,))
    op.Range('B49:B61').Value2 = ((0,), (24,), (.7,), (128,), (512,), (120,), (100,), (200,), (30000,), ('gmres',), ('auto',), (.2,), (1e-10,))
    for row in range(6, 8 if 'JOIN' in kind else 7):
        book.Worksheets('領域').Range(f'K{row}:M{row}').Value2 = ((div, div, div),)
    return op

def compare_solution(app, book, name, mixed=False):
    status = good(app, book, 'API_Solve'); assert str(status).startswith('OK:'), status
    state = json.loads(run(app, book, 'QA_State'))
    model = json.loads(run(app, book, 'ModelJSON')); d = model['dimension']
    desired = {s['name']: s['operator']['backend'] for s in state['regions']}
    def choose(mesh, mat, **settings):
        # Match the independently selected VBA operators, rather than Python timing.
        backend = 'fmm' if mixed and mesh.ne > 30 else 'dense' if mixed else 'fmm'
        op = Dense(mesh, mat) if backend == 'dense' else FMM(mesh, mat, p=settings['p'], leaf=settings['leaf'], theta=settings['theta'], cache_mb=settings['cache_mb'])
        return op, {f'{backend}_estimated_MB': 0, 'selected': backend}
    with threadpool_limits(limits=1), patch('elasticbem.model.choose_operator', choose):
        regions, uu, tt, linear = solve_model(model)
    assert linear['converged'], linear
    n = sum(len(r['mesh'].x) for r in regions)
    rows = np.array(book.Worksheets('境界結果').Range(f'A6:M{5+n}').Value2, object)
    offset = 0; maxu = 0; maxt = 0
    for r, u, t in zip(regions, uu, tt):
        nf = len(u); actual = rows[offset:offset+nf]; offset += nf
        assert all(x == r['name'] for x in actual[:, 0])
        gotu, gott = np.array(actual[:, 7:7+d], float), np.array(actual[:, 10:10+d], float)
        np.testing.assert_allclose(gotu, u, rtol=5e-6, atol=1e-9)
        np.testing.assert_allclose(gott, t, rtol=5e-6, atol=3e-5)
        maxu = max(maxu, float(np.max(abs(gotu-u)))); maxt = max(maxt, float(np.max(abs(gott-t))))
        assert desired[r['name']] == r['op'].stats()['backend']
    record(name, dict(DOF=n*d, backends=desired, residual=state['residual'], max_u_difference_m=maxu, max_t_difference_kPa=maxt))
    return state

def main():
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False; app.DisplayAlerts = False; app.AutomationSecurity = 1; app.EnableEvents = False
    book = None
    try:
        book = app.Workbooks.Open(str(BOOK), 0, True); app.EnableEvents = True
        comp = book.VBProject.VBComponents.Add(1); comp.Name = 'QA_Fmm'
        comp.CodeModule.AddFromString((QA/'QA_Fmm.vba').read_text(encoding='utf-8') + EXTRA_QA)
        op = reset(app, book, '2D_JOIN', 8); op.Range('B13').Value2 = 'auto'; op.Range('B34').Value2 = 200
        book.Worksheets('領域').Range('K6:L6').Value2 = ((4, 8),)
        book.Worksheets('領域').Range('K7:L7').Value2 = ((16, 8),)
        state = compare_solution(app, book, 'mixed dense/FMM bonded two-material regions', mixed=True)
        assert [s['operator']['backend'] for s in state['regions']] == ['dense', 'fmm']
        assert state['regions'][1]['dense_arrays_empty'] and state['regions'][1]['operator']['far_pairs'] > 0

        op = reset(app, book, div=12)
        book.Worksheets('領域').Range('D6:E6').Value2 = ((17.25, -23.5),)
        book.Worksheets('評価点').Range('C6:D6').Value2 = ((17.75, -23),)
        # Loads with coupled components exercise all mixed-BC map columns.
        book.Worksheets('境界条件').Range('H6').Value2 = 0
        book.Worksheets('境界条件').Range('D9:F9').Value2 = ((-67, 't', 23),)
        compare_solution(app, book, 'translated geometry with non-axial traction')

        for state, nu in [('plane_stress', -.25), ('plane_strain', .49)]:
            op = reset(app, book, div=8); op.Range('B6').Value2 = state
            book.Worksheets('材料').Range('B6:C6').Value2 = ((1750, nu),)
            compare_solution(app, book, f'FMM material range {state} nu={nu}')

        op = reset(app, book, '3D_BOX', 1)
        book.Worksheets('材料').Range('C6').Value2 = -.15
        book.Worksheets('境界条件').Range('D10:H10').Value2 = ((-43, 't', 11, 't', 17),)
        compare_solution(app, book, '3D auxetic material with three traction components')

        op = reset(app, book)
        book.Worksheets('評価点').Range('A6:E20').ClearContents()
        good(app, book, 'API_Solve')
        folder = QA/'再検証_評価点なし'; folder.mkdir(exist_ok=True)
        good(app, book, 'API_Export', str(folder))
        report = json.loads((folder/'report.json').read_text(encoding='utf-8'))
        assert report['all_requested_checks_passed'] and report['regions'][0]['interior_reports'] == []
        assert not (folder/'block_interior.csv').exists()
        with np.load(folder/'solution.npz', allow_pickle=False) as data:
            assert not any('_interior_' in k for k in data.files)
        good(app, book, 'API_Import', str(folder))
        record('boundary-only solve and Unicode-folder export/import', 'CSV/NPZ/report valid without interior points')

        op = reset(app, book, '2D_JOIN', 4)
        # Worksheet order differs from region order; NPZ must preserve each point list.
        book.Worksheets('評価点').Range('A6:E9').Value2 = (
            ('right', 'Q2', 1.2, .3, 0), ('left', 'Q1', .2, .3, 0),
            ('right', 'Q1', 1.7, .6, 0), ('left', 'Q2', .7, .6, 0))
        good(app, book, 'API_Solve'); folder = QA/'regression_interleaved'; folder.mkdir(exist_ok=True)
        good(app, book, 'API_Export', str(folder))
        with np.load(folder/'solution.npz', allow_pickle=False) as data:
            np.testing.assert_allclose(data['region_0_interior_points'], [[.2, .3], [.7, .6]])
            np.testing.assert_allclose(data['region_1_interior_points'], [[1.2, .3], [1.7, .6]])
            assert data['region_0_interior_u'].shape == (2, 2) and data['region_1_interior_u'].shape == (2, 2)
        report = json.loads((folder/'report.json').read_text(encoding='utf-8'))
        assert [x['point_id'] for x in report['regions'][0]['interior_reports']] == ['Q1', 'Q2']
        assert [x['point_id'] for x in report['regions'][1]['interior_reports']] == ['Q2', 'Q1']
        good(app, book, 'API_Import', str(folder))
        record('interleaved multi-region point IDs and NPZ ordering', 'two points per region; IDs and values retained')

        op = reset(app, book)
        for cell, changed in [('B13', 'dense'), ('B49', 4), ('B50', 8), ('B51', .8), ('B52', 0),
                              ('B53', 256), ('B54', 20), ('B55', 50), ('B56', 50), ('B57', 2000),
                              ('B58', 'lu'), ('B59', 'gauss'), ('B60', .1), ('B61', 1e-9)]:
            op = reset(app, book); good(app, book, 'API_Solve')
            op.Range(cell).Value2 = changed
            assert run(app, book, 'QA_Invalidated') and '旧入力' in op.Range('B22').Value2
            record('solver input invalidates saved solution '+cell, 'solution and field reports cleared; stale result labelled')

        op = reset(app, book); good(app, book, 'API_Solve')
        folder = QA/'regression_archival'; folder.mkdir(exist_ok=True)
        good(app, book, 'API_Export', str(folder))
        op.Range('B49').Value2 = 4; good(app, book, 'API_Export', str(folder))
        assert not (folder/'report.json').exists() and not (folder/'solution.npz').exists()
        assert any((p/'report.json').exists() and (p/'solution.npz').exists() for p in folder.glob('previous_results_*'))
        assert str(run(app, book, 'API_Import', str(folder))).startswith('ERROR:')
        record('changed FMM input archives full old solution', 'old NPZ/VTK/CSV/report archived; stale import rejected')

        good(app, book, 'API_Solve')
        op.Range('B51').Value2 = 1
        rejected = str(run(app, book, 'API_Solve')); assert rejected.startswith('ERROR:')
        assert not run(app, book, 'QA_BusyState') and app.EnableEvents and app.ScreenUpdating
        op.Range('B51').Value2 = .7; assert str(good(app, book, 'API_Solve')).startswith('OK:')
        folder = QA/'regression_no_folder'; assert not folder.exists()
        assert str(run(app, book, 'API_Export', str(folder))).startswith('ERROR:')
        assert not run(app, book, 'QA_BusyState') and app.EnableEvents
        record('error recovery preserves application state and allows retry', 'invalid theta and nonexistent output folder rejected; retry solved')
    finally:
        app.EnableEvents = False
        if book is not None: book.Close(False)
        app.Quit()
    result = dict(verified_on='2026-10-08', all_passed=True, test_count=len(tests), tests=tests)
    (OUT/'再確認検証結果.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print('ALL PASSED', len(tests), flush=True)

if __name__ == '__main__': main()
