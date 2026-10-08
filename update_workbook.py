"""Embed UTF-8 VBA source updates in the existing native Excel workbook."""
from pathlib import Path
import sys, shutil, time, threading
import pythoncom, win32com.client, win32gui, win32process, win32con

ROOT = Path(__file__).resolve().parent
OUT = ROOT/'outputs/bem_vba_stage3'; QA = ROOT/'qa_stage3'
BOOK = OUT/'Elastic_BEM_VBA_Full.xlsm'
sys.stdout.reconfigure(encoding='utf-8')

def compile_project(app, book):
    errors = []; stop = threading.Event(); pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    def watch():
        while not stop.wait(.1):
            def visit(hwnd, _):
                if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or win32gui.GetClassName(hwnd) != '#32770': return
                texts = []; buttons = []
                def child(h, _):
                    texts.append(win32gui.GetWindowText(h))
                    if win32gui.GetWindowText(h) == 'OK': buttons.append(h)
                win32gui.EnumChildWindows(hwnd, child, None)
                if buttons:
                    errors.append('\n'.join(texts)); win32gui.PostMessage(buttons[0], win32con.BM_CLICK, 0, 0)
            win32gui.EnumWindows(visit, None)
    thread = threading.Thread(target=watch, daemon=True); thread.start()
    try:
        app.VBE.ActiveVBProject = book.VBProject
        book.VBProject.VBComponents('BEM_Solver').CodeModule.CodePane.Show()
        button = app.VBE.CommandBars.FindControl(1, 578)
        if button and button.Enabled: button.Execute()
        for _ in range(20): pythoncom.PumpWaitingMessages(); time.sleep(.05)
        app.VBE.MainWindow.Visible = False
    finally:
        stop.set(); thread.join(2)
    if errors:
        pane = app.VBE.ActiveCodePane; selected = pane.GetSelection()
        print('COMPILE ERROR', errors, pane.CodeModule.Parent.Name, selected, pane.CodeModule.Lines(selected[0], 1), flush=True)
        raise RuntimeError('VBA compilation failed')

def main():
    QA.mkdir(exist_ok=True)
    backup = QA/'pre_issue1.xlsm'
    if not backup.exists(): shutil.copy2(BOOK, backup)
    draft = QA/'updated.xlsm'
    for name in ['vba', 'vba_utf8']: (OUT/name).mkdir(exist_ok=True)
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False; app.DisplayAlerts = False; app.EnableEvents = False; app.AutomationSecurity = 1
    book = None
    try:
        book = app.Workbooks.Open(str(backup), 0, True)
        for path in sorted((ROOT/'vba_src').iterdir()):
            source = path.read_text(encoding='utf-8')
            if path.name == 'ThisWorkbook.vba':
                comp = book.VBProject.VBComponents(book.CodeName)
            else:
                try: book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(path.stem))
                except Exception: pass
                comp = book.VBProject.VBComponents.Add(2 if path.suffix == '.cls' else 1); comp.Name = path.stem
            if comp.CodeModule.CountOfLines: comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(source[source.index('Option Explicit'):])
            export = OUT/'vba'/(path.stem+('.cls' if path.suffix == '.cls' else '.bas'))
            if export.exists(): export.unlink()
            comp.Export(str(export))
            (OUT/'vba_utf8'/path.name).write_bytes(path.read_bytes())
        book.SaveAs(str(draft), 52); book.Close(False); book = app.Workbooks.Open(str(draft))
        compile_project(app, book); app.EnableEvents = True
        for macro, args in [('API_Generate', (True,)), ('API_Solve', ())]:
            result = str(app.Run(f"'{book.Name}'!{macro}", *args)); print(macro, result, flush=True)
            assert result.startswith('OK:'), result
        book.Worksheets('操作').Activate(); book.Save(); book.Close(False); book = None
        shutil.copy2(draft, BOOK)
        print('Updated', BOOK, flush=True)
    finally:
        if book is not None: book.Close(False)
        app.Quit()

if __name__ == '__main__': main()
