from pathlib import Path
import json, zipfile, sys, re, difflib
import numpy as np
import win32com.client

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'outputs/bem_vba_stage3'
QA = ROOT / 'qa_stage3'
BOOK = OUT / 'Elastic_BEM_VBA_Full.xlsm'
sys.stdout.reconfigure(encoding='utf-8')
items = []
seen = set()
reports = ['FMM検証結果.json', 'FMM大規模検証結果.json', 'FMM3D二次要素検証結果.json',
           '演算選択・出力検証結果.json', '境界近傍・入出力検証結果.json', '再確認検証結果.json', 'issue1_regression.json']
for filename in reports:
    data = json.loads((OUT / filename).read_text(encoding='utf-8'))
    assert data['all_passed']
    for item in data['tests']:
        assert item['status'] == 'PASS'
        if item['test'] not in seen:
            items.append(dict(item, source_report=filename)); seen.add(item['test'])

def record(name, detail):
    assert name not in seen
    items.append(dict(test=name, status='PASS', detail=detail)); seen.add(name)
    print('PASS', name, detail, flush=True)

def lexical(value):
    numeric = r'(?:\d+(?:\.\d*)?|\.\d+)(?:[eEdD][+-]?\d+)?[!#%&@]?'
    tokens = re.findall(r'"(?:[^"\n]|"")*"|' + numeric + r'|\w+|[^\s\w]', value)
    result = []
    for token in tokens:
        if re.fullmatch(numeric, token):
            suffix = token[-1] if token[-1] in '!#%&@' else ''
            result.append(('number', float(token.rstrip('!#%&@').replace('D', 'e').replace('d', 'e')), suffix.replace('#', '')))
        else:
            result.append(token if token.startswith('"') else token.lower())
    return result

def run(app, book, name, *args):
    return app.Run(f"'{book.Name}'!{name}", *args)

def good(app, book, name, *args):
    result = run(app, book, name, *args)
    assert not str(result).startswith('ERROR:'), (name, result)
    return result

def main():
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False; app.DisplayAlerts = False; app.AutomationSecurity = 1; app.EnableEvents = False
    book = None
    try:
        book = app.Workbooks.Open(str(BOOK), 0, True); app.EnableEvents = True
        op = book.Worksheets('操作')
        assert book.FileFormat == 52
        assert [s.Name for s in book.Worksheets] == ['操作', 'メッシュ表示', '境界結果', '内点結果', '材料', '領域', '輪郭2D', '境界条件', '接合', '評価点', '節点', '要素', '使い方', 'ログ']
        bindings = []
        for shape in op.Shapes:
            assert shape.Left >= op.Columns(4).Left - .1
            assert shape.Left + shape.Width <= op.Columns(4).Left + op.Columns(4).Width + .1
            assert str(shape.OnAction).startswith('BEM_')
            bindings.append(str(shape.OnAction))
        assert len(bindings) == 8 and 'BEM_Solve' in bindings
        assert op.Range('B13').Value2 == 'fmm' and op.Range('B58').Value2 == 'gmres'
        for cell in ['B13', 'B58', 'B59']:
            assert op.Range(cell).Validation.Type == 3
        record('final FMM workbook native sheets, buttons and selectors', {'sheets': 14, 'buttons': 8, 'macro_format': 52})

        count = 0
        for p in (ROOT / 'vba_src').glob('*'):
            component = book.VBProject.VBComponents(book.CodeName if p.name == 'ThisWorkbook.vba' else p.stem)
            code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines).replace('\r\n', '\n').strip()
            expected = p.read_text(encoding='utf-8'); expected = expected[expected.index('Option Explicit'):].strip()
            a, b = lexical(code), lexical(expected)
            equivalent = len(a) == len(b) and all(one == two or (isinstance(one, tuple) and isinstance(two, tuple) and one[2] == two[2] and np.isclose(one[1], two[1], rtol=1e-14, atol=0)) for one, two in zip(a, b))
            if not equivalent:
                print('\n'.join(list(difflib.unified_diff(expected.splitlines(), code.splitlines()))[:30]), flush=True)
                raise AssertionError(p.name)
            assert not re.search(r'(?im)^\s*(?:Call\s+)?(?:Shell\b|RunPython\b)|\.Run\s+"(?:python|py )', code)
            count += 1
        assert not any(c.Name.startswith('QA_') for c in book.VBProject.VBComponents)
        record('final embedded VBA sources including FMM match delivered sources', {'source_count': count, 'temporary_QA_modules': 0, 'Python_execution': False})

        row = book.Worksheets('内点結果').Range('A6:Q6').Value2[0]
        np.testing.assert_allclose(row[5:7], [.0015166666666666666, -.00065], atol=1e-8, rtol=0)
        np.testing.assert_allclose(row[8:14], [100, 0, 30, 0, 0, 0], atol=.001, rtol=0)
        assert row[14] and row[16] and '収束OK' in op.Range('B22').Value2
        assert [op.Range(f'B{i}').Value2 for i in range(18, 22)] == [128, 64, 192, 384]
        record('saved FMM initial sample solution', {'DOF': 384, 'ux_m': row[5], 'uy_m': row[6], 'sxx_kPa': row[8], 'szz_kPa': row[10]})

        status = good(app, book, 'API_Solve'); assert str(status).startswith('OK: 384 DOF')
        folder = OUT / 'example'; folder.mkdir(exist_ok=True)
        # The final package exports a clean example through the embedded VBA only.
        for p in folder.iterdir():
            if p.is_file(): p.unlink()
        good(app, book, 'API_Export', str(folder))
        report = json.loads((folder / 'report.json').read_text(encoding='utf-8'))
        linear = report['linear_solver']; operator = report['regions'][0]['operator']
        assert report['all_requested_checks_passed'] and linear['converged']
        assert linear['backend'] == 'fmm' and linear['method'] == 'gmres' and linear['DOF'] == 384
        assert operator['far_pairs'] == 60 and operator['backend'] == 'fmm'
        assert 'FMM' in report['accuracy_scope']
        with np.load(folder / 'solution.npz', allow_pickle=False) as data:
            assert int(data['dimension']) == 2
            assert data['region_0_u'].shape == (192, 2)
            np.testing.assert_allclose(data['region_0_interior_stress'][0], [[100, 0, 0], [0, 0, 0], [0, 0, 30]], atol=.001, rtol=0)
            assert json.loads(str(data['model_json']))['solver']['backend'] == 'fmm'
        assert (folder / 'block_boundary.vtk').exists()
        record('final workbook reopened, solved and exported entirely in VBA', {'DOF': linear['DOF'], 'true_relative_residual': linear['true_relative_residual'], 'far_pairs': operator['far_pairs'], 'formats': 'JSON, CSV, VTK, NPZ'})

        chart = book.Worksheets('メッシュ表示').ChartObjects(1).Chart
        assert chart.SeriesCollection().Count == 5 and chart.SeriesCollection(5).Points().Count == 128
        chart.Export(str(QA / 'native_final_2D.png'))
        record('final native mesh chart retained', '128 geometric node markers and 4 boundary series')

        good(app, book, 'API_Sample', '2D_RECT')
        op.Range('B13').Value2 = 'fmm'; op.Range('B58').Value2 = 'gmres'
        book.Worksheets('領域').Range('K6:L6').Value2 = ((4, 4),)
        bc = book.Worksheets('境界条件'); bc.Range('A6:H25').ClearContents()
        bc.Range('A6:H6').Value2 = (('block', 'ALL', 'u', 0, 'u', 0, 'u', 0),)
        status = good(app, book, 'API_Solve'); assert status == 'OK: 96 DOF / residual=0'
        row = book.Worksheets('内点結果').Range('A6:Q6').Value2[0]
        np.testing.assert_equal(np.array(row[5:7] + row[8:14], float), 0)
        np.testing.assert_equal(np.array(book.Worksheets('境界結果').Range('H6:I53').Value2, float), 0)
        np.testing.assert_equal(np.array(book.Worksheets('境界結果').Range('K6:L53').Value2, float), 0)
        record('FMM GMRES exact zero-RHS solution', 'all displacements, tractions and internal stresses zero; residual=0')
        version = str(app.Version)
    finally:
        app.EnableEvents = False
        if book is not None: book.Close(False)
        app.Quit()

    inspection = (QA / 'visual_inspection.ndjson').read_text(encoding='utf-8')
    assert 'Cell search matched 0 entries.' in inspection
    for name in ['操作', '解析設定', 'FMM設定', '境界結果', '内点結果', '使い方']:
        assert (QA / f'final_{name}.png').exists()
    record('final spreadsheet render and formula error scan', 'six rendered sheets/ranges inspected; no spreadsheet error cells')
    with zipfile.ZipFile(BOOK) as z:
        assert z.testzip() is None and 'xl/vbaProject.bin' in z.namelist()
        assert b'macroEnabled' in z.read('[Content_Types].xml')
    record('final macro-enabled XLSM package integrity', 'native xl/vbaProject.bin present; all ZIP CRCs valid')
    result = dict(verified_on='2026-10-09', excel_version=version, pass_count=len(items), all_passed=True, tests=items)
    (OUT / '検証結果.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8', newline='\n')
    package = ROOT / 'outputs/Elastic_BEM_VBA_Stage3.zip'
    with zipfile.ZipFile(package, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in sorted(OUT.rglob('*')):
            if p.is_file(): z.write(p, p.relative_to(OUT))
    with zipfile.ZipFile(package) as z: assert z.testzip() is None
    print(json.dumps(dict(pass_count=len(items), book_bytes=BOOK.stat().st_size, package_bytes=package.stat().st_size, package=str(package)), ensure_ascii=False), flush=True)

if __name__ == '__main__': main()
