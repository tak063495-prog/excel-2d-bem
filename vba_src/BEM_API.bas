Attribute VB_Name = "BEM_API"
Option Explicit

'Callable entry points return errors as text for external automation.
Public Function API_Solve() As String
    On Error GoTo Failed
    API_Solve = SolveCore(): Exit Function
Failed:
    API_Solve = "ERROR: " & Err.description
End Function
Public Function API_Generate(Optional ByVal preview As Boolean = True) As String
    On Error GoTo Failed
    API_Generate = GenerateCore(preview): Exit Function
Failed:
    API_Generate = "ERROR: " & Err.description
End Function
Public Function API_Export(ByVal folder As String) As String
    On Error GoTo Failed
    API_Export = ExportCore(folder): Exit Function
Failed:
    API_Export = "ERROR: " & Err.description
End Function
Public Function API_Import(ByVal folder As String) As String
    On Error GoTo Failed
    API_Import = ImportCore(folder): Exit Function
Failed:
    API_Import = "ERROR: " & Err.description
End Function
Public Function API_Sample(ByVal kind As String) As String
    On Error GoTo Failed
    LoadSample kind: API_Sample = "OK": Exit Function
Failed:
    API_Sample = "ERROR: " & Err.description
End Function
