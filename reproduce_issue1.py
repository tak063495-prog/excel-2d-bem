from pathlib import Path
import json, sys, re
import win32com.client

ROOT = Path(__file__).resolve().parent
if len(sys.argv) != 2:
    raise SystemExit('Usage: python reproduce_issue1.py PATH_TO_UNMODIFIED_3510d4c_WORKBOOK')
BOOK = Path(sys.argv[1]).resolve()
QA = ROOT/'qa_stage3'; QA.mkdir(exist_ok=True)
sys.stdout.reconfigure(encoding='utf-8')
results = {}

def run(app, book, name, *args): return app.Run(f"'{book.Name}'!{name}", *args)

def setup(app, book, offset):
    assert run(app, book, 'API_Sample', '2D_RECT') == 'OK'
    book.Worksheets('領域').Range('D6:E6').Value2 = ((offset, offset),)
    book.Worksheets('領域').Range('K6:L6').Value2 = ((4, 4),)
    book.Worksheets('評価点').Range('C6:D6').Value2 = ((offset+.5, offset+.5),)
    book.Worksheets('操作').Range('B13').Value2 = 'dense'
    book.Worksheets('操作').Range('B58').Value2 = 'gmres'

def main():
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False; app.DisplayAlerts = False; app.AutomationSecurity = 1; app.EnableEvents = False
    book = None
    try:
        book = app.Workbooks.Open(str(BOOK), 0, True); app.EnableEvents = True
        setup(app, book, 0)
        results['original_origin0_generate'] = run(app, book, 'API_Generate', False)
        setup(app, book, 1e8)
        results['original_offset1e8_generate'] = run(app, book, 'API_Generate', False)
        module = book.VBProject.VBComponents('BEM_Mesh').CodeModule
        code = module.Lines(1, module.CountOfLines)
        code, count = re.subn(r'q\s*=\s*r\.Vertices\(a\s*\+\s*1\)\s*:\s*c\s*=\s*r\.Vertices\(b\s*\+\s*1\)', 'q = Sub3(r.Vertices(a + 1), r.Origin): c = Sub3(r.Vertices(b + 1), r.Origin)', code, flags=re.I)
        assert count == 1, count
        module.DeleteLines(1, module.CountOfLines); module.AddFromString(code)
        setup(app, book, 1e8)
        results['area_only_offset1e8_generate'] = run(app, book, 'API_Generate', False)
        results['area_only_offset1e8_solve'] = run(app, book, 'API_Solve')
        qa = book.VBProject.VBComponents.Add(1); qa.Name = 'QA_Number'
        qa.CodeModule.AddFromString('''Option Explicit
Public Function QA_Formats() As Variant
Dim value As Double
value = 100000000# + 1# / 96#
QA_Formats = Array(value, JNum(value), Format$(value, "0.0000000000000000E+00"), CStr(CDec(value)))
End Function''')
        results['number_formats'] = run(app, book, 'QA_Formats')
    finally:
        (QA/'issue1_reproduction.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(results, ensure_ascii=False, indent=2), flush=True)
        app.EnableEvents = False
        if book is not None: book.Close(False)
        app.Quit()

if __name__ == '__main__': main()
