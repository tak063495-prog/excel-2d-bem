"""Native Excel regression for issue #1; leaves the delivered workbook unchanged."""
from pathlib import Path
import json, struct
import numpy as np
import win32com.client
from verify_fmm import ROOT, OUT, QA, BOOK, run, good

TEST_VBA = '''
Option Explicit
Public Function QA_NumberText() As Variant
    Dim values As Variant, result() As Variant, i As Long, count As Long
    count = ThisWorkbook.Worksheets("NumberQA").Cells(Rows.Count, 1).End(xlUp).Row
    values = ThisWorkbook.Worksheets("NumberQA").Range("A1:A" & count).Value2
    ReDim result(1 To count, 1 To 3)
    For i = 1 To count
        result(i, 1) = CDbl(values(i, 1)): result(i, 2) = JNum(CDbl(values(i, 1)))
        result(i, 3) = (Val(result(i, 2)) = CDbl(values(i, 1)))
    Next i
    QA_NumberText = result
End Function
Public Function QA_SelfDistances(ByVal name As String) As Variant
    Dim sr As CSolveRegion, g As CElementData, rule As CIntegrationRule
    Dim f As Long, e As Long, k As Long, j As Long, dist As Double, minimum As Double, zero As Long
    Dim p(1 To 3) As Double, n(1 To 3) As Double, ds(1 To 3) As Double, dt(1 To 3) As Double, jac As Double
    Set sr = SolveData(name): minimum = 1E+99
    For f = 1 To sr.NF
        e = (f - 1) \\ sr.Ref.FieldCount
        Set g = sr.Geometries(e + 1)
        Set rule = SelfRule(sr.Ref.Dimension, sr.Ref.ElementOrder, (f - 1) Mod sr.Ref.FieldCount + 1)
        For k = 1 To rule.Count
            MapGeometry g, rule.A(k), rule.B(k), p, n, jac, ds, dt: dist = 0
            For j = 1 To sr.Ref.Dimension: dist = dist + (p(j) - sr.X(f, j)) ^ 2: Next j
            If dist = 0 Then zero = zero + 1
            If dist < minimum Then minimum = dist
        Next k
    Next f
    QA_SelfDistances = Array(zero, Sqr(minimum))
End Function
Public Function QA_ReverseTopology() As String
    Dim r As CBemRegion, reversed As Collection, ids As Variant, out As Variant, i As Long, j As Long
    On Error GoTo Rejected
    Set r = Models("block"): Set reversed = New Collection
    For i = 1 To r.Elements.Count
        ids = r.Elements(i): out = ids
        For j = 0 To UBound(ids): out(j) = ids(UBound(ids) - j): Next j
        reversed.Add out
    Next i
    Set r.Elements = reversed: ValidateTopology r: QA_ReverseTopology = "ACCEPTED": Exit Function
Rejected:
    QA_ReverseTopology = Err.Description
End Function
Public Function QA_DefaultFrame() As Variant
    Dim sr As CSolveRegion: Set sr = New CSolveRegion
    QA_DefaultFrame = sr.FrameOrigin
End Function
'''
tests = []

def record(name, **detail):
    tests.append(dict(test=name, status='PASS', detail=detail))
    print('PASS', name, detail, flush=True)

def configure(app, book, kind, order, offset, backend='dense', level=3, div=4):
    good(app, book, 'API_Sample', kind)
    op = book.Worksheets('操作'); reg = book.Worksheets('領域')
    op.Range('B7').Value2 = order; op.Range('B13').Value2 = backend
    op.Range('B37').Value2 = level; op.Range('B43').Value2 = 1e-10
    op.Range('B58').Value2 = 'gmres'; op.Range('B49').Value2 = 6 if kind.startswith('2') else 4
    op.Range('B50').Value2 = 8
    d = int(kind[0]); count = 2 if 'JOIN' in kind else 1
    for row in range(6, 6+count):
        old = reg.Range(f'D{row}:F{row}').Value2[0]
        reg.Range(f'D{row}:F{row}').Value2 = (tuple((old[j] or 0)+offset[j] for j in range(3)),)
        reg.Range(f'K{row}:M{row}').Value2 = ((div, div, div),)
    pts = book.Worksheets('評価点')
    for row in range(6, 6+count):
        old = pts.Range(f'C{row}:E{row}').Value2[0]
        pts.Range(f'C{row}:E{row}').Value2 = (tuple((old[j] or 0)+offset[j] for j in range(3)),)
    status = good(app, book, 'API_Generate', False)
    status = good(app, book, 'API_Solve'); assert str(status).startswith('OK:'), status
    state = json.loads(run(app, book, 'QA_State'))
    n = int(op.Range('B21').Value2/d)
    rows = np.array(book.Worksheets('境界結果').Range(f'A6:M{5+n}').Value2, object)
    internal = np.array(book.Worksheets('内点結果').Range(f'A6:Q{5+count}').Value2, object)
    assert np.isfinite(np.array(rows[:, 7:7+d], float)).all()
    return rows, internal, state

def main():
    app = win32com.client.DispatchEx('Excel.Application'); app.Visible = False
    app.DisplayAlerts = False; app.AutomationSecurity = 1; app.EnableEvents = False; book = None
    try:
        book = app.Workbooks.Open(str(BOOK), 0, True); app.EnableEvents = True
        comp = book.VBProject.VBComponents.Add(1); comp.Name = 'QA_Issue1'
        comp.CodeModule.AddFromString((QA/'QA_Fmm.vba').read_text(encoding='utf-8') + TEST_VBA.replace('Option Explicit', ''))
        assert tuple(run(app, book, 'QA_DefaultFrame')) == (0., 0., 0.)
        record('standalone solver region initializes coordinate frame (runtime error 13)')
        sheet = book.Worksheets.Add(); sheet.Name = 'NumberQA'
        rng = np.random.default_rng(164)
        values = np.r_[0, 1e8+1/96, -(1e8+1/96), np.nextafter(1e8, np.inf),
                       np.pi, 1e-300, 1e300, np.nextafter(0., 1.), np.finfo(float).max,
                       rng.uniform(-1e8, 1e8, 1500),
                       np.ldexp(rng.uniform(.5, 1., 1500), rng.integers(-1000, 1000, 1500))]
        sheet.Range(f'A1:A{len(values)}').Value2 = tuple((float(x),) for x in values)
        # Compare the actual COM input double too: Excel may flush subnormal values.
        actual = run(app, book, 'QA_NumberText')
        for number, text, equal in actual:
            assert struct.pack('d', float(text)) == struct.pack('d', number), (number, text)
            assert equal, (number, text, 'VBA Val')
        app.UseSystemSeparators = False; app.DecimalSeparator = ','; app.ThousandsSeparator = '.'
        for number, text, equal in run(app, book, 'QA_NumberText'):
            assert float(text) == number and equal and ',' not in text, (number, text)
        app.UseSystemSeparators = True
        record('double text round-trip and decimal-comma locale', numbers=len(values))

        shift = (1e8, 1e8, 0.)
        for backend, order, level, div in [('dense', 0, 0, 4), ('dense', 2, 3, 4),
                                          ('dense', 2, 4, 4), ('fmm', 2, 3, 16)]:
            base, point0, _ = configure(app, book, '2D_RECT', order, (0., 0., 0.), backend, level, div)
            action0 = np.array(run(app, book, 'QA_Apply', 'block', 1.2), float)
            moved, point1, state = configure(app, book, '2D_RECT', order, shift, backend, level, div)
            action1 = np.array(run(app, book, 'QA_Apply', 'block', 1.2), float)
            np.testing.assert_allclose(action1, action0, rtol=1e-12, atol=1e-13)
            for col in [7, 10]: np.testing.assert_allclose(np.array(moved[:, col:col+2], float), np.array(base[:, col:col+2], float), rtol=1e-10, atol=1e-10)
            np.testing.assert_allclose(np.array(point1[:, 8:14], float), np.array(point0[:, 8:14], float), atol=1e-8)
            zero, distance = run(app, book, 'QA_SelfDistances', 'block'); assert zero == 0 and distance > 0
            if backend == 'fmm': assert state['regions'][0]['operator']['far_pairs'] > 0
            record(f'large-offset 2D {backend} order={order} DE={level}', min_self_distance_m=distance, residual=state['residual'], operator_max_difference=float(np.max(abs(action1-action0))))

        folder = QA/'issue1_large_export'; folder.mkdir(exist_ok=True)
        good(app, book, 'API_Export', str(folder)); model = json.loads((folder/'model.json').read_text(encoding='utf-8'))
        csv = np.loadtxt(folder/'block_boundary.csv', delimiter=',', skiprows=1)
        np.testing.assert_array_equal(csv[:, :2], np.array(moved[:, 4:6], float))
        with np.load(folder/'solution.npz', allow_pickle=False) as data:
            np.testing.assert_array_equal(data['region_0_collocation'], csv[:, :2])
            np.testing.assert_array_equal(data['region_0_vertices'], model['regions'][0]['vertices'])
        vtk = (folder/'block_boundary.vtk').read_text(encoding='utf-8').splitlines()
        start = next(i for i, line in enumerate(vtk) if line.startswith('POINTS '))
        n = int(vtk[start].split()[1]); points = np.array([[float(x) for x in line.split()] for line in vtk[start+1:start+1+n]])
        vertices = np.array(model['regions'][0]['vertices'])
        expected = vertices[np.array(model['regions'][0]['elements']).ravel()]
        np.testing.assert_array_equal(points[:, :2], expected)
        good(app, book, 'API_Import', str(folder))
        record('world coordinates retained in JSON/CSV/NPZ/VTK and CSV import', vertex_count=n)

        for kind, order, div, origin in [('2D_JOIN', 2, 4, (1e8-2, -1e8+2, 0.)), ('3D_BOX', 0, 1, (1e8, -1e8, 1e8)),
                                        ('3D_BOX', 2, 1, (1e8, -1e8, 1e8)), ('3D_JOIN', 0, 1, (1e8-2, -1e8+2, 1e8-2))]:
            d = int(kind[0]); base, _, _ = configure(app, book, kind, order, (0., 0., 0.), div=div)
            moved, _, state = configure(app, book, kind, order, origin, div=div)
            for col in [7, 10]: np.testing.assert_allclose(np.array(moved[:, col:col+d], float), np.array(base[:, col:col+d], float), rtol=2e-9, atol=1e-8)
            record('translated '+kind+' order='+str(order), residual=state['residual'])
        configure(app, book, '2D_RECT', 2, shift)
        rejected = run(app, book, 'QA_ReverseTopology'); assert '反時計回り' in rejected, rejected
        record('clockwise topology still rejected', message=rejected)
        good(app, book, 'API_Sample', '2D_RECT'); book.Worksheets('領域').Range('G6').Value2 = 0
        assert str(run(app, book, 'API_Generate', False)).startswith('ERROR:')
        record('zero-size geometry still rejected')
        for clockwise in [False, True]:
            good(app, book, 'API_Sample', '2D_POLYGON')
            op = book.Worksheets('操作'); op.Range('B13').Value2 = 'dense'
            book.Worksheets('領域').Range('K6').Value2 = 4
            contour = book.Worksheets('輪郭2D'); contour.Range('A6:D20').ClearContents()
            vertices = [(1e8, 1e8), (1e8+1, 1e8), (1e8+1, 1e8+1), (1e8, 1e8+1)]
            if clockwise: vertices.reverse()
            contour.Range('A6:D9').Value2 = tuple(('block', i+1, *p) for i, p in enumerate(vertices))
            book.Worksheets('評価点').Range('C6:D6').Value2 = ((1e8+.5, 1e8+.5),)
            assert str(good(app, book, 'API_Solve')).startswith('OK:')
            record('large polygon '+('clockwise input normalized' if clockwise else 'counterclockwise input'))
        contour.Range('A6:D9').Value2 = tuple(('block', i+1, 1e8+x, 1e8+y) for i, (x, y) in enumerate([(0, 0), (1, 0), (1, 2**-20), (0, 2**-20)]))
        book.Worksheets('評価点').Range('C6:D6').Value2 = ((1e8+.5, 1e8+2**-21),)
        good(app, book, 'API_Generate', False)
        record('thin nondegenerate large-offset polygon accepted', height_m=2**-20)
    finally:
        if book is not None: book.Close(False)
        app.Quit()
    result = dict(verified_on='2026-10-09', all_passed=True, test_count=len(tests), tests=tests)
    (OUT/'issue1_regression.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print('ALL PASSED', len(tests), flush=True)

if __name__ == '__main__': main()
