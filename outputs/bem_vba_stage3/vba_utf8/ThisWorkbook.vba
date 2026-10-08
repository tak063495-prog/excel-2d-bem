Option Explicit

Private Sub Workbook_Open()
    ApplyDimension
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Busy Then Exit Sub
    Select Case Sh.Name
    Case "操作"
        If Intersect(Target, Sh.Range("B5:B13,B34:B43,B49:B61")) Is Nothing Then Exit Sub
        MarkChanged
        If Not Intersect(Target, Sh.Range("B5")) Is Nothing Then ApplyDimension
    Case "材料", "領域", "輪郭2D", "境界条件", "接合", "評価点"
        If Target.row + Target.Rows.Count - 1 < FIRST_ROW Then Exit Sub
        MarkChanged
    End Select
End Sub
