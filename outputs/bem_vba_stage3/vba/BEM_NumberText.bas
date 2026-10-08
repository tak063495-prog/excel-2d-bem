Attribute VB_Name = "BEM_NumberText"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Sub CopyNumberBits Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal count As LongPtr)
#Else
Private Declare Sub CopyNumberBits Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal count As Long)
#End If
Public Function RoundTripNumber(ByVal value As Double) As String
    'Str$ alone keeps about 15 digits, which can lose metre offsets at 1E8.
    'Recover the exact binary significand, then scale in Decimal (28 digits).
    Dim fast As String, bytes(0 To 7) As Byte, exponent As Long, binaryPower As Long, decimalPower As Long
    Dim high As Long, low As Double, mantissa As Variant, lower As Variant, upper As Variant
    Dim digits As String, scientificPower As Long, point As Long, i As Long, negative As Boolean
    On Error GoTo Precise
    fast = Trim$(Str$(value))
    If left$(fast, 1) = "." Then fast = "0" & fast
    If left$(fast, 2) = "-." Then fast = "-0" & mid$(fast, 2)
    If Val(fast) = value Then RoundTripNumber = fast: Exit Function
Precise:
    On Error GoTo 0
    CopyNumberBits bytes(0), value, 8
    negative = (bytes(7) And &H80) <> 0
    exponent = (CLng(bytes(7) And &H7F) * 16) + (bytes(6) \ 16)
    If exponent = 2047 Then Fail "JSON/CSVÇ…óLå¿êîà»äOÇÕèëÇ´çûÇﬂÇ‹ÇπÇÒÅB"
    high = CLng(bytes(6) And &HF) * 65536 + CLng(bytes(5)) * 256 + bytes(4)
    low = CDbl(bytes(3)) * 16777216# + CDbl(bytes(2)) * 65536# + CDbl(bytes(1)) * 256# + bytes(0)
    If exponent = 0 Then binaryPower = -1074 Else high = high + 1048576: binaryPower = exponent - 1075
    mantissa = CDec(high) * CDec(4294967296#) + CDec(low)
    If mantissa = 0 Then RoundTripNumber = "0": Exit Function
    lower = CDec(1E+16): upper = CDec(1E+17)
    Do While mantissa < lower
        mantissa = mantissa * CDec(10): decimalPower = decimalPower - 1
    Loop
    If binaryPower >= 0 Then
        For i = 1 To binaryPower
            mantissa = mantissa * CDec(2)
            If mantissa >= upper Then mantissa = mantissa / CDec(10): decimalPower = decimalPower + 1
        Next i
    Else
        For i = 1 To -binaryPower
            mantissa = mantissa / CDec(2)
            If mantissa < lower Then mantissa = mantissa * CDec(10): decimalPower = decimalPower - 1
        Next i
    End If
    mantissa = Fix(mantissa + CDec(0.5))
    If mantissa >= upper Then mantissa = mantissa / CDec(10): decimalPower = decimalPower + 1
    digits = CStr(mantissa): scientificPower = decimalPower + 16
    Do While Len(digits) > 1 And right$(digits, 1) = "0": digits = left$(digits, Len(digits) - 1): Loop
    If scientificPower >= -4 And scientificPower <= 15 Then
        point = scientificPower + 1
        If point <= 0 Then
            digits = "0." & String$(-point, "0") & digits
        ElseIf point >= Len(digits) Then
            digits = digits & String$(point - Len(digits), "0")
        Else
            digits = left$(digits, point) & "." & mid$(digits, point + 1)
        End If
    Else
        If Len(digits) > 1 Then digits = left$(digits, 1) & "." & mid$(digits, 2)
        digits = digits & "E" & CStr(scientificPower)
    End If
    If negative Then digits = "-" & digits
    RoundTripNumber = digits
End Function

