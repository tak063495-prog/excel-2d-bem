VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ApplyDimension
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Busy Then Exit Sub
    Select Case Sh.Name
    Case "‘€ì"
        If Intersect(Target, Sh.Range("B5:B13,B34:B43,B49:B61")) Is Nothing Then Exit Sub
        MarkChanged
        If Not Intersect(Target, Sh.Range("B5")) Is Nothing Then ApplyDimension
    Case "Ş—¿", "—Ìˆæ", "—ÖŠs2D", "‹«ŠEğŒ", "Ú‡", "•]‰¿“_"
        If Target.row + Target.Rows.Count - 1 < FIRST_ROW Then Exit Sub
        MarkChanged
    End Select
End Sub

