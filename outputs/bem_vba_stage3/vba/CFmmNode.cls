VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFmmNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Number As Long, Level As Long, Parent As Long
Public Half0 As Double, HalfWidth As Double
Public Ids As Variant, Center As Variant, Knots As Variant
Public Transfer As Variant, SourceT As Variant, SourceUN As Variant, Target As Variant
Public Children As Collection
Public NearSources As Variant
Private Sub Class_Initialize()
    Set Children = New Collection
End Sub

