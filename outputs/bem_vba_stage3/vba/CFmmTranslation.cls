VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFmmTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Key As String
Public TargetNode As Long, SourceNode As Long
Public Targets As Collection, Sources As Collection
Public Coefficients As Variant
Public LastUse As Long
Private Sub Class_Initialize()
    Set Targets = New Collection: Set Sources = New Collection
End Sub

