Attribute VB_Name = "BinOctHexImpl"
Option Explicit

' Convert an integer to binary string with "0b" prefix.
Public Function PyBinCore(ByVal value As Variant) As String
    If Not IsObject(value) Then
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyBinCore = LongToBinStr(CLng(value))
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyBin", "argument must be an integer"
        End Select
        Exit Function
    End If

    ' Guard: object must implement IIndex
    If Not TypeOf value Is IIndex Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyBin", "object does not implement IIndex"
        Exit Function
    End If

    Dim indexImpl As IIndex: Set indexImpl = value
    PyBinCore = LongToBinStr(indexImpl.Index())
End Function

' Convert an integer to octal string with "0o" prefix.
Public Function PyOctCore(ByVal value As Variant) As String
    ' Fast path: numeric integers
    If Not IsObject(value) Then
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyOctCore = LongToOctStr(CLng(value))
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyOct", "argument must be an integer"
        End Select
        Exit Function
    End If

    ' Guard: object must implement IIndex
    If Not TypeOf value Is IIndex Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyOct", "object does not implement IIndex"
        Exit Function
    End If

    Dim indexImpl As IIndex: Set indexImpl = value
    PyOctCore = LongToOctStr(indexImpl.Index())
End Function

' Convert an integer to lowercase hex string with "0x" prefix.
Public Function PyHexCore(ByVal value As Variant) As String
    ' Fast path: numeric integers
    If Not IsObject(value) Then
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyHexCore = LongToHexStr(CLng(value))
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyHex", "argument must be an integer"
        End Select
        Exit Function
    End If

    ' Guard: object must implement IIndex
    If Not TypeOf value Is IIndex Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyHex", "object does not implement IIndex"
        Exit Function
    End If

    Dim indexImpl As IIndex: Set indexImpl = value
    PyHexCore = LongToHexStr(indexImpl.Index())
End Function

' ─── Private helpers ───

Private Function LongToBinStr(ByVal raw As Long) As String
    If raw = 0 Then
        LongToBinStr = "0b0"
        Exit Function
    End If
    If raw = &H80000000 Then
        LongToBinStr = "-0b10000000000000000000000000000000"
        Exit Function
    End If
    If raw < 0 Then
        LongToBinStr = "-0b" & PosToBinStr(-raw)
        Exit Function
    End If
    LongToBinStr = "0b" & PosToBinStr(raw)
End Function

Private Function PosToBinStr(ByVal val As Long) As String
    Dim buf As String
    Dim pos As Long

    buf = Space$(31)  ' max 31 bits for 32-bit Long (sign bit handled by caller)

    pos = 31
    Do While val > 0
        Mid$(buf, pos) = CStr(val And 1)
        pos = pos - 1
        val = val \ 2
    Loop

    If pos = 31 Then
        PosToBinStr = ""
        Exit Function
    End If
    PosToBinStr = Mid$(buf, pos + 1)
End Function

Private Function LongToOctStr(ByVal raw As Long) As String
    If raw = 0 Then
        LongToOctStr = "0o0"
        Exit Function
    End If
    If raw = &H80000000 Then
        LongToOctStr = "-0o20000000000"
        Exit Function
    End If
    If raw < 0 Then
        LongToOctStr = "-0o" & PosToOctStr(-raw)
        Exit Function
    End If
    LongToOctStr = "0o" & PosToOctStr(raw)
End Function

Private Function PosToOctStr(ByVal val As Long) As String
    Dim buf As String
    Dim pos As Long

    buf = Space$(11)  ' max 11 octal digits for 32-bit

    pos = 11
    Do While val > 0
        Mid$(buf, pos) = CStr(val Mod 8)
        pos = pos - 1
        val = val \ 8
    Loop

    If pos = 11 Then
        PosToOctStr = ""
        Exit Function
    End If
    PosToOctStr = Mid$(buf, pos + 1)
End Function

Private Function LongToHexStr(ByVal raw As Long) As String
    If raw = 0 Then
        LongToHexStr = "0x0"
        Exit Function
    End If
    If raw = &H80000000 Then
        LongToHexStr = "-0x80000000"
        Exit Function
    End If
    If raw < 0 Then
        LongToHexStr = "-0x" & PosToHexStr(-raw)
        Exit Function
    End If
    LongToHexStr = "0x" & PosToHexStr(raw)
End Function

Private Function PosToHexStr(ByVal val As Long) As String
    Dim buf As String
    Dim pos As Long

    buf = Space$(8)  ' max 8 hex digits for 32-bit

    pos = 8
    Do While val > 0
        Mid$(buf, pos) = Mid$("0123456789abcdef", (val And &HF) + 1, 1)
        pos = pos - 1
        val = val \ 16
    Loop

    If pos = 8 Then
        PosToHexStr = ""
        Exit Function
    End If
    PosToHexStr = Mid$(buf, pos + 1)
End Function