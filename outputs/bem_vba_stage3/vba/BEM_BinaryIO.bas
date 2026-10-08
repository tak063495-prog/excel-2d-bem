Attribute VB_Name = "BEM_BinaryIO"
Option Explicit
Private CRCTable(0 To 255) As Long, CRCReady As Boolean
Private Function ShiftRight(ByVal value As Long, ByVal bits As Long) As Long
    ShiftRight = (value And &H7FFFFFFF) \ CLng(2 ^ bits)
    If value < 0 Then ShiftRight = ShiftRight Or CLng(2 ^ (31 - bits))
End Function
Private Function CRC32(ByRef bytes() As Byte) As Long
    Dim i As Long, j As Long, value As Long, crcValue As Long
    If Not CRCReady Then
        For i = 0 To 255
            value = i
            For j = 1 To 8
                If (value And 1) <> 0 Then value = ShiftRight(value, 1) Xor &HEDB88320 Else value = ShiftRight(value, 1)
            Next j
            CRCTable(i) = value
        Next i
        CRCReady = True
    End If
    crcValue = &HFFFFFFFF
    For i = LBound(bytes) To UBound(bytes): crcValue = ShiftRight(crcValue, 8) Xor CRCTable((crcValue Xor CLng(bytes(i))) And &HFF): Next i
    CRC32 = Not crcValue
End Function
Private Function NPYBuffer(ByVal dtype As String, ByVal shape As String) As CByteBuffer
    Dim buffer As CByteBuffer, header As String, padding As Long
    Set buffer = New CByteBuffer
    header = "{'descr': '" & dtype & "', 'fortran_order': False, 'shape': " & shape & ", }"
    padding = (64 - ((10 + Len(header) + 1) Mod 64)) Mod 64: header = header & Space$(padding) & vbLf
    buffer.AppendByte 147: buffer.AppendASCII "NUMPY": buffer.AppendByte 1: buffer.AppendByte 0
    buffer.AppendUInt16 Len(header): buffer.AppendASCII header: Set NPYBuffer = buffer
End Function
Private Sub AddEntry(ByVal entries As Collection, ByVal name As String, ByVal buffer As CByteBuffer)
    Dim entry As CZipEntry, content() As Byte
    Set entry = New CZipEntry: entry.name = name & ".npy": content = buffer.bytes: entry.content = content: entry.size = UBound(content) + 1: entry.CRC = CRC32(content): entries.Add entry
End Sub
Private Sub WriteZIP(ByVal entries As Collection, ByVal path As String)
    Dim buffer As CByteBuffer, entry As CZipEntry, content() As Byte, centralStart As Long, centralSize As Long
    Set buffer = New CByteBuffer
    If entries.Count > 65535 Then Fail "NPZの配列数がZIP上限を超えます。"
    For Each entry In entries
        entry.Offset = buffer.Count
        buffer.AppendUInt32 &H4034B50: buffer.AppendUInt16 20: buffer.AppendUInt16 0: buffer.AppendUInt16 0
        buffer.AppendUInt16 0: buffer.AppendUInt16 33: buffer.AppendUInt32 CDbl(entry.CRC)
        buffer.AppendUInt32 CDbl(entry.size): buffer.AppendUInt32 CDbl(entry.size): buffer.AppendUInt16 Len(entry.name): buffer.AppendUInt16 0
        buffer.AppendASCII entry.name: content = entry.content: buffer.AppendBytes content
    Next entry
    centralStart = buffer.Count
    For Each entry In entries
        buffer.AppendUInt32 &H2014B50: buffer.AppendUInt16 20: buffer.AppendUInt16 20: buffer.AppendUInt16 0: buffer.AppendUInt16 0
        buffer.AppendUInt16 0: buffer.AppendUInt16 33: buffer.AppendUInt32 CDbl(entry.CRC): buffer.AppendUInt32 CDbl(entry.size): buffer.AppendUInt32 CDbl(entry.size)
        buffer.AppendUInt16 Len(entry.name): buffer.AppendUInt16 0: buffer.AppendUInt16 0: buffer.AppendUInt16 0: buffer.AppendUInt16 0
        buffer.AppendUInt32 0: buffer.AppendUInt32 CDbl(entry.Offset): buffer.AppendASCII entry.name
    Next entry
    centralSize = buffer.Count - centralStart
    buffer.AppendUInt32 &H6054B50: buffer.AppendUInt16 0: buffer.AppendUInt16 0: buffer.AppendUInt16 entries.Count: buffer.AppendUInt16 entries.Count
    buffer.AppendUInt32 CDbl(centralSize): buffer.AppendUInt32 CDbl(centralStart): buffer.AppendUInt16 0: buffer.SaveFile path
End Sub
Public Sub SaveSolutionNPZ(ByVal folder As String)
    Dim entries As Collection, buffer As CByteBuffer, model As String, name As Variant, sr As CSolveRegion, r As CBemRegion, prefix As String
    Dim region As Long, i As Long, j As Long, k As Long, d As Long, point As Variant, ids As Variant, xyz() As Double, u() As Double, t() As Double
    Dim interior As Variant, last As Long, row As Long, available As Boolean, stress(1 To 3, 1 To 3) As Double, fso As Object
    Set entries = New Collection: model = ModelJSON(): Set buffer = NPYBuffer("<U" & Len(model), "()")
    buffer.AppendUTF32 model: AddEntry entries, "model_json", buffer
    Set buffer = NPYBuffer("<i8", "()"): buffer.AppendInt64 CurrentDimension(): AddEntry entries, "dimension", buffer
    last = WS("内点結果").Cells(WS("内点結果").rows.Count, 1).End(xlUp).row
    If last >= FIRST_ROW Then interior = WS("内点結果").Range("A6:Q" & last).Value2
    For Each name In SolveData.keys
        Set sr = SolveData(name): Set r = sr.Ref: d = r.Dimension: prefix = "region_" & region & "_": region = region + 1
        Set buffer = NPYBuffer("<f8", "(" & r.Vertices.Count & ", " & d & ")")
        For Each point In r.Vertices: For j = 0 To d - 1: buffer.AppendDouble CDbl(point(j)): Next j: Next point
        AddEntry entries, prefix & "vertices", buffer
        ids = r.elements(1): Set buffer = NPYBuffer("<i8", "(" & r.elements.Count & ", " & UBound(ids) + 1 & ")")
        For Each ids In r.elements: For j = 0 To UBound(ids): buffer.AppendInt64 CLng(ids(j)): Next j: Next ids
        AddEntry entries, prefix & "elements", buffer
        xyz = sr.X: u = sr.u: t = sr.t
        Set buffer = NPYBuffer("<f8", "(" & sr.NF & ", " & d & ")")
        For i = 1 To sr.NF: For j = 1 To d: buffer.AppendDouble xyz(i, j): Next j: Next i
        AddEntry entries, prefix & "collocation", buffer
        Set buffer = NPYBuffer("<f8", "(" & sr.NF & ", " & d & ")")
        For i = 1 To sr.ND: buffer.AppendDouble u(i): Next i
        AddEntry entries, prefix & "u", buffer
        Set buffer = NPYBuffer("<f8", "(" & sr.NF & ", " & d & ")")
        For i = 1 To sr.ND: buffer.AppendDouble t(i): Next i
        AddEntry entries, prefix & "t", buffer
        If r.Points.Count > 0 Then
            Set buffer = NPYBuffer("<f8", "(" & r.Points.Count & ", " & d & ")")
            For Each point In r.Points: For j = 0 To d - 1: buffer.AppendDouble CDbl(point(j)): Next j: Next point
            AddEntry entries, prefix & "interior_points", buffer
            Set buffer = NPYBuffer("<f8", "(" & r.Points.Count & ", " & d & ")")
            For row = 1 To UBound(interior, 1)
                If CStr(interior(row, 1)) = CStr(name) Then
                    available = LCase$(CStr(interior(row, 17))) = "true"
                    For j = 1 To d: If available Then buffer.AppendDouble CDbl(interior(row, 5 + j)) Else buffer.AppendNaN
                    Next j
                End If
            Next row
            AddEntry entries, prefix & "interior_u", buffer
            Set buffer = NPYBuffer("<f8", "(" & r.Points.Count & ", 3, 3)")
            For row = 1 To UBound(interior, 1)
                If CStr(interior(row, 1)) = CStr(name) Then
                    available = LCase$(CStr(interior(row, 17))) = "true"
                    If available Then
                        stress(1, 1) = interior(row, 9): stress(2, 2) = interior(row, 10): stress(3, 3) = interior(row, 11)
                        stress(1, 2) = interior(row, 12): stress(2, 1) = stress(1, 2): stress(2, 3) = interior(row, 13): stress(3, 2) = stress(2, 3): stress(3, 1) = interior(row, 14): stress(1, 3) = stress(3, 1)
                    End If
                    For j = 1 To 3: For k = 1 To 3
                        If available Then buffer.AppendDouble stress(j, k) Else buffer.AppendNaN
                    Next k: Next j
                End If
            Next row
            AddEntry entries, prefix & "interior_stress", buffer
        End If
    Next name
    Set fso = CreateObject("Scripting.FileSystemObject"): WriteZIP entries, fso.BuildPath(folder, "solution.npz")
End Sub

