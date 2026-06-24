Attribute VB_Name = "IntCoreImpl"
Option Explicit

' Core implementation for PyInt — returns an integer matching Python's int().
Public Function PyIntCore(Optional ByVal value As Variant, Optional ByVal base As Variant) As Long
    If IsArray(value) Then
        Ensure.IsTrue False, CommonHResults.TypeMismatch, _
            "Builtin.PyInt", "argument must be a number or string, not an array"

    ElseIf IsMissing(value) Then
        PyIntCore = 0

    ElseIf Not IsMissing(base) Then
        PyIntCore = ParseIntString(CStr(value), AsIntegerIndex(base))

    ElseIf IsObject(value) Then
        If TypeOf value Is IIndex Then
            Dim idx As IIndex: Set idx = value
            PyIntCore = Builtin.PyInt(idx.Index())
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyInt", "object does not implement IIndex"
        End If

    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyIntCore = value
            Case vbSingle, vbDouble, vbCurrency, vbDecimal
                PyIntCore = Fix(value)
            Case vbString
                PyIntCore = ParseIntString(value, 10)
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyInt", "argument must be a number or string"
        End Select
    End If
End Function

' Validates and converts a value to a Long integer index.
' Only allows Byte, Integer, Long, and objects implementing IIndex.
' Raises error for other types.
Public Function AsIntegerIndex(ByVal value As Variant) As Long
    If IsObject(value) Then
        ' Use PyInt for IIndex protocol — recursion handled internally
        AsIntegerIndex = Builtin.PyInt(value)
    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                AsIntegerIndex = CLng(value)
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyRange", "argument must be an integer"
        End Select
    End If
End Function

' Converts a single character to its digit value (0-35).
' Returns -1 for non-digit characters.
Public Function CharToDigit(ByVal c As String) As Long
    Dim code As Long: code = Asc(c)
    If code >= 48 And code <= 57 Then           ' 0-9
        CharToDigit = code - 48
    ElseIf code >= 65 And code <= 90 Then       ' A-Z
        CharToDigit = code - 65 + 10
    ElseIf code >= 97 And code <= 122 Then      ' a-z
        CharToDigit = code - 97 + 10
    Else
        CharToDigit = -1
    End If
End Function

' Parses a string as an integer in the given base (2-36).
' Accepts optional 0x/0o/0b prefix when it matches the base or base=0.
' Raises Ensure on empty string, invalid digits, or overflow.
' TODO: full whitespace stripping + underscore grouping (Python 3.6+)
Public Function ParseIntString(ByVal s As String, ByVal base As Long) As Long
    s = Trim(s)

    If Len(s) = 0 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.ParseIntString", "empty string"
    End If

    Dim pos As Long: pos = 1
    Dim negative As Boolean: negative = False

    Select Case Mid$(s, pos, 1)
        Case "-": negative = True: pos = pos + 1
        Case "+": pos = pos + 1
    End Select

    ' Detect 0x/0o/0b prefix when base=0 or base matches
    Dim actualBase As Long: actualBase = base
    If pos <= Len(s) - 2 Then
        If Mid$(s, pos, 1) = "0" Then
            Dim p As String: p = UCase$(Mid$(s, pos + 1, 1))
            If p = "X" And (base = 0 Or base = 16) Then
                actualBase = 16: pos = pos + 2
            ElseIf p = "O" And (base = 0 Or base = 8) Then
                actualBase = 8: pos = pos + 2
            ElseIf p = "B" And (base = 0 Or base = 2) Then
                actualBase = 2: pos = pos + 2
            End If
        End If
    End If

    If base = 0 And actualBase = 0 Then actualBase = 10

    If actualBase < 2 Or actualBase > 36 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.ParseIntString", "invalid base"
    End If

    ' Parse digits using Double for intermediate storage (handles Long.MinValue)
    Dim accum As Double: accum = 0
    Dim c As String
    Dim digitVal As Long

    Do While pos <= Len(s)
        c = Mid$(s, pos, 1): pos = pos + 1
        ' TODO: skip underscores (_) here

        digitVal = CharToDigit(c)
        If digitVal < 0 Or digitVal >= actualBase Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.ParseIntString", "invalid digit for base " & CStr(actualBase)
        End If

        accum = accum * actualBase + digitVal

        ' Fast overflow detection
        If accum > 2147483648# Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.ParseIntString", "integer overflow"
        End If
    Loop

    If negative Then
        If accum > 2147483648# Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.ParseIntString", "integer overflow"
        End If
        ParseIntString = CLng(-accum)
    Else
        If accum > 2147483647# Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.ParseIntString", "integer overflow"
        End If
        ParseIntString = CLng(accum)
    End If
End Function