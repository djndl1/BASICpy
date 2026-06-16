VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Builtin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private m_binDigits() As String
Private m_octDigits() As String
Private m_hexDigits() As String

Private Sub Class_Initialize()
    ReDim m_octDigits(0 To 7)
    m_octDigits(0) = "0"
    m_octDigits(1) = "1"
    m_octDigits(2) = "2"
    m_octDigits(3) = "3"
    m_octDigits(4) = "4"
    m_octDigits(5) = "5"
    m_octDigits(6) = "6"
    m_octDigits(7) = "7"

    ReDim m_hexDigits(0 To 15)
    m_hexDigits(0) = "0"
    m_hexDigits(1) = "1"
    m_hexDigits(2) = "2"
    m_hexDigits(3) = "3"
    m_hexDigits(4) = "4"
    m_hexDigits(5) = "5"
    m_hexDigits(6) = "6"
    m_hexDigits(7) = "7"
    m_hexDigits(8) = "8"
    m_hexDigits(9) = "9"
    m_hexDigits(10) = "a"
    m_hexDigits(11) = "b"
    m_hexDigits(12) = "c"
    m_hexDigits(13) = "d"
    m_hexDigits(14) = "e"
    m_hexDigits(15) = "f"

    ReDim m_binDigits(0 To 1)
    m_binDigits(0) = "0"
    m_binDigits(1) = "1"
End Sub

Public Function PyAll(ByRef arr As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If

    If IsArray(arr) Then
        PyAll = PyAllArray(arr, predicate)
    ElseIf TypeOf arr Is Collection Then
        PyAll = PyAllCollection(arr, predicate)
    ElseIf IsObject(arr) And IsIEnumVARAINT(arr) Then
        PyAll = PyAllIEnumVariant(arr, predicate)
    Else
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyAll", "argument is neither an array nor a Collection nor an enumerable object"
    End If
End Function

Public Function PyAny(ByRef arr As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If

    If IsArray(arr) Then
        PyAny = PyAnyArray(arr, predicate)
    ElseIf TypeName(arr) = "Collection" Then
        PyAny = PyAnyCollection(arr, predicate)
    ElseIf IsObject(arr) And IsIEnumVARAINT(arr) Then
        PyAny = PyAnyIEnumVariant(arr, predicate)
    Else
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyAny", "argument is neither an array nor a Collection nor an enumerable object"
    End If
End Function

Public Function PyAbs(ByVal value As Variant) As Variant
    If IsObject(value) Then
        If TypeOf value Is IAbsolute Then
            Dim absImpl As IAbsolute
            Set absImpl = value
            PyAbs = absImpl.Absolute()
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyAbs", "object does not implement IAbsolute"
        End If
    Else
        PyAbs = Abs(value)
    End If
End Function

Public Function PyBin(ByVal value As Variant) As String
    If IsObject(value) Then
        If TypeOf value Is IIndex Then
            Dim indexImpl As IIndex
            Set indexImpl = value
            PyBin = PyBin(indexImpl.Index())
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyBin", "object does not implement IIndex"
        End If
    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyBin = LongToBinStr(CLng(value))
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyBin", "argument must be an integer"
        End Select
    End If
End Function

Public Function PyOct(ByVal value As Variant) As String
Attribute PyOct.VB_Description = "Convert an integer to an octal string, prefixed with ""0o"" (or ""-0o"" for negative values).\r\nArgs:\r\n  value (Variant): Integer value to convert (Byte, Integer, or Long).\r\nReturns:\r\n  String: Octal representation, e.g. oct(8) -> ""0o10"", oct(-56) -> ""-0o70""."
    If IsObject(value) Then
        If TypeOf value Is IIndex Then
            Dim indexImpl As IIndex
            Set indexImpl = value
            PyOct = PyOct(indexImpl.Index())
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyOct", "object does not implement IIndex"
        End If
    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyOct = LongToOctStr(CLng(value))
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyOct", "argument must be an integer"
        End Select
    End If
End Function

Public Function PyHex(ByVal value As Variant) As String
Attribute PyHex.VB_Description = "Convert an integer to a lowercase hexadecimal string, prefixed with ""0x"" (or ""-0x"" for negative values).\r\nArgs:\r\n  value (Variant): Integer value to convert (Byte, Integer, or Long).\r\nReturns:\r\n  String: Hex representation, e.g. hex(255) -> ""0xff"", hex(-42) -> ""-0x2a""."
    If IsObject(value) Then
        If TypeOf value Is IIndex Then
            Dim indexImpl As IIndex
            Set indexImpl = value
            PyHex = PyHex(indexImpl.Index())
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyHex", "object does not implement IIndex"
        End If
    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyHex = LongToHexStr(CLng(value))
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyHex", "argument must be an integer"
        End Select
    End If
End Function

' Returns the length of a value, dispatching across VB6 types.
' Arrays use UBound-LBound+1, strings use Len(s), objects dispatch through
' ILength (interface protocol), TypeOf (intrinsic controls), TypeName (COM class
' names), or a dynamic Count/Length fallback. Non-len-able types raise an error.
' Args:
'   value (Variant): The value whose length to return. Can be an array, string,
'     object implementing ILength, VB6 intrinsic control (ListBox, Form, etc.),
'     COM collection (Dictionary, ArrayList, etc.), or any COM object with
'     a Count/Length property.
' Returns:
'   Long: The number of elements or characters in the value.
Public Function PyLen(ByVal value As Variant) As Long
Attribute PyLen.VB_Description = "Returns the length of a value, dispatching across VB6 types.\r\nArrays use UBound-LBound+1, strings use Len(s), objects dispatch through ILength (interface protocol), TypeOf (intrinsic controls), TypeName (COM class names), or a dynamic Count/Length fallback. Non-len-able types raise an error.\r\nArgs:\r\n  value (Variant): The value whose length to return. Can be an array, string, object implementing ILength, VB6 intrinsic control (ListBox, Form, etc.), COM collection (Dictionary, ArrayList, etc.), or any COM object with a Count/Length property.\r\nReturns:\r\n  Long: The number of elements or characters in the value."
    If IsArray(value) Then
        PyLen = UBound(value) - LBound(value) + 1
    ElseIf VarType(value) = vbString Then
        PyLen = Len(value)
    ElseIf IsObject(value) Then
        ' TypeOf checks for built-in types and ILength implementations.
        ' Unrecognized objects fall through to PyLenByName, which uses TypeName
        ' dispatch for COM types without a compile-time reference.
        If TypeOf value Is ILength Then
            Dim lenImpl As ILength
            Set lenImpl = value
            PyLen = lenImpl.Length()
        Else
            Dim matched As Boolean
            PyLen = PyLenByTypeOf(value, matched)
            If Not matched Then
                PyLen = PyLenByName(value)
            End If
        End If
    Else
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyLen", "argument has no len()"
    End If
End Function

' Dispatches to Count/Length/RecordCount based on TypeName(value).
' Matches COM classes by their runtime class name string, avoiding the need for
' a compile-time reference to their type library. Extend by adding Case labels.
' Args:
'   value (Variant): The object whose TypeName is matched against known classes.
' Returns:
'   Long: The object's Count, Length, RecordCount, or raises an error.
Private Function PyLenByName(ByVal value As Variant) As Long
    Select Case TypeName(value)
        Case "Dictionary", "ICollection", "ArrayList", "BitArray", "Hashtable", "SortedList", "Stack", "Queue"
            PyLenByName = value.Count
        Case "CSlice"
            PyLenByName = value.Length
        Case "Recordset"
            PyLenByName = value.RecordCount
        Case Else
            ' Dynamic fallback: try Count and Length via IDispatch late binding.
            On Error Resume Next
            PyLenByName = value.Count
            If Err.Number = 0 Then
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear
            PyLenByName = value.Length
            If Err.Number = 0 Then
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyLen", "argument has no len()"
    End Select
End Function

' Matches intrinsic VB6 controls (Collection, ListBox, Form, etc.) via TypeOf.
' The TypeOf operator uses a vtable check (faster than TypeName string comparison).
' Each control type maps to a different property: .Count for Collection/Form,
' .ListCount for ListBox/ComboBox/DirListBox/FileListBox.
' Args:
'   value (Variant): The object to check against known VB6 intrinsic types.
'   matched (Boolean): Set to True if a TypeOf match was found, False otherwise.
' Returns:
'   Long: The Count or ListCount of the matched control, or 0 if unmatched.
Private Function PyLenByTypeOf(ByVal value As Variant, ByRef matched As Boolean) As Long
    matched = True
    If TypeOf value Is Collection Then
        PyLenByTypeOf = value.Count
    ElseIf TypeOf value Is ListBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is ComboBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is DirListBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is FileListBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is Form Then
        PyLenByTypeOf = value.Count
    ElseIf TypeOf value Is MDIForm Then
        PyLenByTypeOf = value.Count
    ElseIf TypeOf value Is PropertyPage Then
        PyLenByTypeOf = value.Count
    Else
        matched = False
        PyLenByTypeOf = 0
    End If
End Function

Private Function PyAllArray(ByRef arr As Variant, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Not predicate.Test(arr(i)) Then
            PyAllArray = False
            Exit Function
        End If
    Next i
    PyAllArray = True
End Function

Private Function PyAllCollection(ByVal col As Collection, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If Not predicate.Test(col(i)) Then
            PyAllCollection = False
            Exit Function
        End If
    Next i
    PyAllCollection = True
End Function

Private Function PyAllIEnumVariant(ByVal obj As Object, ByVal predicate As IPredicate) As Boolean
    Dim element As Variant
    For Each element In obj
        If Not predicate.Test(element) Then
            PyAllIEnumVariant = False
            Exit Function
        End If
    Next element
    PyAllIEnumVariant = True
End Function

Private Function PyAnyArray(ByRef arr As Variant, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If predicate.Test(arr(i)) Then
            PyAnyArray = True
            Exit Function
        End If
    Next i
    PyAnyArray = False
End Function

Private Function PyAnyCollection(ByVal col As Collection, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If predicate.Test(col(i)) Then
            PyAnyCollection = True
            Exit Function
        End If
    Next i
    PyAnyCollection = False
End Function

Private Function PyAnyIEnumVariant(ByVal obj As Object, ByVal predicate As IPredicate) As Boolean
    Dim element As Variant
    For Each element In obj
        If predicate.Test(element) Then
            PyAnyIEnumVariant = True
            Exit Function
        End If
    Next element
    PyAnyIEnumVariant = False
End Function

Private Function LongToBinStr(ByVal raw As Long) As String
    If raw = 0 Then
        LongToBinStr = "0b0"
        Exit Function
    End If
    If raw < 0 Then
        If raw = &H80000000 Then
            LongToBinStr = "-0b10000000000000000000000000000000"
        Else
            LongToBinStr = "-0b" & PosToBinStr(-raw)
        End If
    Else
        LongToBinStr = "0b" & PosToBinStr(raw)
    End If
End Function

Private Function PosToBinStr(ByVal val As Long) As String
    Dim buf As String
    Dim pos As Long

    buf = Space$(31)  ' max 31 bits for 32-bit Long (sign bit handled by caller)

    pos = 31
    Do While val > 0
        Mid$(buf, pos) = m_binDigits(val And 1)
        pos = pos - 1
        val = val \ 2
    Loop

    If pos = 31 Then
        PosToBinStr = ""
    Else
        PosToBinStr = Mid$(buf, pos + 1)
    End If
End Function

Private Function LongToOctStr(ByVal raw As Long) As String
    If raw = 0 Then
        LongToOctStr = "0o0"
        Exit Function
    End If
    If raw < 0 Then
        If raw = &H80000000 Then
            LongToOctStr = "-0o20000000000"
        Else
            LongToOctStr = "-0o" & PosToOctStr(-raw)
        End If
    Else
        LongToOctStr = "0o" & PosToOctStr(raw)
    End If
End Function

Private Function PosToOctStr(ByVal val As Long) As String
    Dim buf As String
    Dim pos As Long

    buf = Space$(11)  ' max 11 octal digits for 32-bit

    pos = 11
    Do While val > 0
        Mid$(buf, pos) = m_octDigits(val Mod 8)
        pos = pos - 1
        val = val \ 8
    Loop

    If pos = 11 Then
        PosToOctStr = ""
    Else
        PosToOctStr = Mid$(buf, pos + 1)
    End If
End Function

Private Function LongToHexStr(ByVal raw As Long) As String
    If raw = 0 Then
        LongToHexStr = "0x0"
        Exit Function
    End If
    If raw < 0 Then
        If raw = &H80000000 Then
            LongToHexStr = "-0x80000000"
        Else
            LongToHexStr = "-0x" & PosToHexStr(-raw)
        End If
    Else
        LongToHexStr = "0x" & PosToHexStr(raw)
    End If
End Function

Private Function PosToHexStr(ByVal val As Long) As String
    Dim buf As String
    Dim pos As Long

    buf = Space$(8)  ' max 8 hex digits for 32-bit

    pos = 8
    Do While val > 0
        Mid$(buf, pos) = m_hexDigits(val And &HF)
        pos = pos - 1
        val = val \ 16
    Loop

    If pos = 8 Then
        PosToHexStr = ""
    Else
        PosToHexStr = Mid$(buf, pos + 1)
    End If
End Function

Private Function IsIEnumVARAINT(ByVal obj As Object) As Boolean
    On Error Resume Next
    Dim enumObj As Object
    Set enumObj = CallByName(obj, "_NewEnum", VbGet)
    If Err.Number <> 0 Then
        IsIEnumVARAINT = False
    Else
        IsIEnumVARAINT = Not (enumObj Is Nothing)
    End If
    On Error GoTo 0
End Function