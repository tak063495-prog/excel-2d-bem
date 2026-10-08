Attribute VB_Name = "BEM_FmmSettings"
Option Explicit
Public RequestedBackend As String, LinearMethod As String
Public FmmOrder As Long, FmmLeaf As Long, MaxFmmUnknowns As Long, ExpectedIterations As Long
Public FmmTheta As Double, FmmCacheMB As Double, MemoryBudgetMB As Double
Public GMRESRestart As Long, GMRESMaxCycles As Long
Public SolveBackend As String, SolveIterations As Long, SolveCycles As Long, SolveLinearInfo As Long
Private Type MemoryStatus
    Length As Long
    Load As Long
    TotalPhysical As Currency
    AvailablePhysical As Currency
    TotalPage As Currency
    AvailablePage As Currency
    TotalVirtual As Currency
    AvailableVirtual As Currency
    Extended As Currency
End Type
#If VBA7 Then
Private Declare PtrSafe Function GlobalMemoryStatusEx Lib "kernel32" (ByRef status As MemoryStatus) As Long
#Else
Private Declare Function GlobalMemoryStatusEx Lib "kernel32" (ByRef status As MemoryStatus) As Long
#End If
Public Function AvailableMemoryMB() As Double
    Dim status As MemoryStatus: status.Length = LenB(status)
    If GlobalMemoryStatusEx(status) <> 0 Then AvailableMemoryMB = CDbl(status.AvailablePhysical) * 10000 / 1048576 Else AvailableMemoryMB = 512
End Function
Public Sub ReadFmmOptions()
    Dim configured As Double
    RequestedBackend = LCase$(TextAt(WS("操作"), 13, 2))
    If RequestedBackend <> "auto" And RequestedBackend <> "dense" And RequestedBackend <> "fmm" Then Fail "演算方式はauto、dense、fmmです。"
    FmmOrder = IntegerSetting(49, 0, 0, IIf(CurrentDimension() = 2, 8, 6))
    If FmmOrder = 0 Then FmmOrder = IIf(CurrentDimension() = 2, 6, 4)
    If FmmOrder < 2 Then Fail "FMM次数pは0（自動）または2D:2～8、3D:2～6です。"
    FmmLeaf = IntegerSetting(50, 24, 1, 10000)
    FmmTheta = SolveSetting(51, 0.7)
    If FmmTheta <= 0 Or FmmTheta >= 1 Then Fail "FMM thetaは0より大きく1より小さい数です。"
    FmmCacheMB = SolveSetting(52, 128): configured = SolveSetting(53, 512)
    If FmmCacheMB < 0 Or FmmCacheMB > 8192 Then Fail "FMMキャッシュは0～8192 MBです。"
    If configured <= 0 Or configured > 16384 Then Fail "モデルのメモリ上限は0より大きく16384 MB以下です。"
    MemoryBudgetMB = MinDouble(configured, 0.6 * AvailableMemoryMB())
    ExpectedIterations = IntegerSetting(54, 120, 1, 100000)
    GMRESRestart = IntegerSetting(55, 100, 1, 1000)
    GMRESMaxCycles = IntegerSetting(56, 200, 1, 10000)
    MaxFmmUnknowns = IntegerSetting(57, 30000, 1, 200000)
    LinearMethod = LCase$(TextAt(WS("操作"), 58, 2)): If LinearMethod = "" Then LinearMethod = "gmres"
    If LinearMethod <> "gmres" And LinearMethod <> "lu" Then Fail "求解法はgmresまたはluです。"
    If RequestedBackend = "fmm" And LinearMethod = "lu" Then Fail "FMMでは求解法gmresを選択してください。"
End Sub
