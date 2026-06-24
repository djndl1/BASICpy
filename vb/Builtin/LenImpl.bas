Attribute VB_Name = "LenImpl"
Option Explicit

' Core implementation: returns the length of a value. Raises if the value has no len().
Public Function PyLenCore(ByRef value As Variant) As Long
    PyLenCore = PyLenImpl(value)
    If PyLenCore = -1 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyLen", "argument has no len()"
    End If
End Function

' Safe internal length — returns -1 on failure instead of raising.
' Preconditions checked before operations — no catch-all On Error needed.
' Used by PyLen (public) and PyBool.
Public Function PyLenImpl(ByRef value As Variant) As Long
    ' Arrays — local OERN for undimensioned check (rare path, cold cost)
    If IsArray(value) Then
        On Error Resume Next
        PyLenImpl = UBound(value) - LBound(value) + 1
        If Err.Number <> 0 Then PyLenImpl = -1
        On Error GoTo 0
        Exit Function
    End If

    ' Strings — Len never fails
    If VarType(value) = vbString Then
        PyLenImpl = Len(value)
        Exit Function
    End If

    If Not IsObject(value) Then
        ' Non-lengthable type
        PyLenImpl = -1
    End If

    ' Nothing — no length, return -1 (caller raises error)
    If value Is Nothing Then
        PyLenImpl = -1
        Exit Function
    End If

    If TypeOf value Is ILength Then
        Dim lenImpl As ILength: Set lenImpl = value
        PyLenImpl = lenImpl.Length()
        Exit Function
    End If

    Dim matched As Boolean: PyLenImpl = PyLenByTypeOf(value, matched)
    If Not matched Then
        PyLenImpl = PyLenByName(value)    ' returns -1 on failure, no raise
    End If
    Exit Function
End Function

' Dispatches to Count/Length/RecordCount based on TypeName(value).
' Matches COM classes by their runtime class name string, avoiding the need for
' a compile-time reference to their type library. Extend by adding Case labels.
' Args:
'   value (Variant): The object whose TypeName is matched against known classes.
' Returns:
'   Long: The object's Count, Length, RecordCount, or -1 if unmatched/dynamic.
Public Function PyLenByName(ByVal value As Variant) As Long
    Select Case TypeName(value)
        Case "Dictionary", "ICollection", "ArrayList", "BitArray", "Hashtable", "SortedList", "Stack", "Queue"
            PyLenByName = value.count
        Case "CSlice"
            PyLenByName = value.Length
        Case "Recordset"
            PyLenByName = value.RecordCount
        Case Else
            ' Dynamic fallback: try Count and Length via IDispatch late binding.
            On Error Resume Next
            PyLenByName = value.count
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
            PyLenByName = -1    ' No length found — signal failure
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
Public Function PyLenByTypeOf(ByVal value As Variant, ByRef matched As Boolean) As Long
    matched = True
    If TypeOf value Is Collection Then
        PyLenByTypeOf = value.count
    ElseIf TypeOf value Is ListBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is ComboBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is DirListBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is FileListBox Then
        PyLenByTypeOf = value.ListCount
    ElseIf TypeOf value Is Form Then
        PyLenByTypeOf = value.count
    ElseIf TypeOf value Is MDIForm Then
        PyLenByTypeOf = value.count
    ElseIf TypeOf value Is PropertyPage Then
        PyLenByTypeOf = value.count
    Else
        matched = False
        PyLenByTypeOf = 0
    End If
End Function