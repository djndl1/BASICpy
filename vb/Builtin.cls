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

Public Function PyLen(ByVal value As Variant) As Long
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

Private Function PyLenByName(ByVal value As Variant) As Long
    ' TypeName dispatch matches COM types by class name without requiring a
    ' compile-time reference to their type library.  Add new types here.
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

Private Function PyLenByTypeOf(ByVal value As Variant, ByRef matched As Boolean) As Long
    ' TypeOf checks for intrinsic VB6 controls — no TypeName or reference needed.
    ' Returns the .ListCount or .Count via late binding on the Variant.
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
    Dim result As String
    result = ""
    Do While val > 0
        If (val And 1) = 1 Then
            result = "1" & result
        Else
            result = "0" & result
        End If
        val = val \ 2
    Loop
    PosToBinStr = result
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