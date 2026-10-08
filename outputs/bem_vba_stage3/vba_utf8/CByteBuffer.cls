VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CByteBuffer"
Option Explicit
Private Data() As Byte, Used As Long, Capacity As Long
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If
Private Sub Reserve(ByVal increment As Long)
    Dim required As Long: required = Used + increment
    If required > Capacity Then
        Capacity = MaxLong(required, MaxLong(256, Capacity * 2)): ReDim Preserve Data(0 To Capacity - 1)
    End If
End Sub
Public Property Get Count() As Long
    Count = Used
End Property
Public Sub AppendByte(ByVal value As Byte)
    Reserve 1: Data(Used) = value: Used = Used + 1
End Sub
Public Sub AppendUInt16(ByVal value As Long)
    Reserve 2: Data(Used) = value Mod 256: Data(Used + 1) = value \ 256: Used = Used + 2
End Sub
Public Sub AppendUInt32(ByVal value As Double)
    Dim i As Long
    If value < 0 Then value = value + 4294967296#
    Reserve 4
    For i = 0 To 3: Data(Used + i) = value - 256 * Fix(value / 256): value = Fix(value / 256): Next i
    Used = Used + 4
End Sub
Public Sub AppendInt64(ByVal value As Long)
    AppendUInt32 CDbl(value): AppendUInt32 IIf(value < 0, 4294967295#, 0#)
End Sub
Public Sub AppendDouble(ByVal value As Double)
    Reserve 8: CopyMemory Data(Used), value, 8: Used = Used + 8
End Sub
Public Sub AppendNaN()
    AppendUInt32 0: AppendUInt32 2146959360#
End Sub
Public Sub AppendASCII(ByVal value As String)
    Dim i As Long
    Reserve Len(value)
    For i = 1 To Len(value): Data(Used) = AscW(Mid$(value, i, 1)): Used = Used + 1: Next i
End Sub
Public Sub AppendUTF32(ByVal value As String)
    Dim i As Long, code As Long
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1)): If code < 0 Then code = code + 65536
        AppendUInt32 CDbl(code)
    Next i
End Sub
Public Sub AppendBytes(ByRef value() As Byte)
    Dim count As Long: count = UBound(value) - LBound(value) + 1
    Reserve count: CopyMemory Data(Used), value(LBound(value)), count: Used = Used + count
End Sub
Public Function Bytes() As Variant
    Dim result() As Byte
    If Used = 0 Then Exit Function
    ReDim result(0 To Used - 1): CopyMemory result(0), Data(0), Used: Bytes = result
End Function
Public Sub SaveFile(ByVal path As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 1: stream.Open: stream.Write Bytes(): stream.SaveToFile path, 2: stream.Close
End Sub
