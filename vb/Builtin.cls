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

Public Function PyAll(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If

    Dim iter As IIterator
    Set iter = PyIter(value)
    PyAll = PyAllFromIter(iter, predicate)
End Function

Public Function PyAny(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If

    Dim iter As IIterator
    Set iter = PyIter(value)
    PyAny = PyAnyFromIter(iter, predicate)
End Function

' Iterates through an IIterator, checking all elements pass the predicate.
Private Function PyAllFromIter(ByVal iterator As IIterator, ByVal predicate As IPredicate) As Boolean
    Dim item As Variant
    item = iterator.NextItem()
    Do While Not IsError(item)
        If Not predicate.Test(item) Then
            PyAllFromIter = False
            Exit Function
        End If
        item = iterator.NextItem()
    Loop
    PyAllFromIter = True
End Function

' Iterates through an IIterator, checking if any element passes the predicate.
Private Function PyAnyFromIter(ByVal iterator As IIterator, ByVal predicate As IPredicate) As Boolean
    Dim item As Variant
    item = iterator.NextItem()
    Do While Not IsError(item)
        If predicate.Test(item) Then
            PyAnyFromIter = True
            Exit Function
        End If
        item = iterator.NextItem()
    Loop
    PyAnyFromIter = False
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
            PyBin = PyBin(indexImpl.index())
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
            PyOct = PyOct(indexImpl.index())
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
            PyHex = PyHex(indexImpl.index())
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
Private Function PyLenImpl(ByRef value As Variant) As Long
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

    ' Objects
    If IsObject(value) Then
        ' Nothing → no length, return -1 (caller raises error)
        If value Is Nothing Then
            PyLenImpl = -1
            Exit Function
        End If
        If TypeOf value Is ILength Then
            Dim lenImpl As ILength
            Set lenImpl = value
            PyLenImpl = lenImpl.Length()
            Exit Function
        Else
            Dim matched As Boolean
            PyLenImpl = PyLenByTypeOf(value, matched)
            If Not matched Then
                PyLenImpl = PyLenByName(value)    ' returns -1 on failure, no raise
            End If
            Exit Function
        End If
    End If

    ' Non-lengthable type
    PyLenImpl = -1
End Function

Public Function PyLen(ByRef value As Variant) As Long
Attribute PyLen.VB_Description = "Returns the length of a value, dispatching across VB6 types.\r\nArrays use UBound-LBound+1, strings use Len(s), objects dispatch through ILength (interface protocol), TypeOf (intrinsic controls), TypeName (COM class names), or a dynamic Count/Length fallback. Non-len-able types raise an error.\r\nArgs:\r\n  value (Variant): The value whose length to return. Can be an array, string, object implementing ILength, VB6 intrinsic control (ListBox, Form, etc.), COM collection (Dictionary, ArrayList, etc.), or any COM object with a Count/Length property.\r\nReturns:\r\n  Long: The number of elements or characters in the value."
    PyLen = PyLenImpl(value)
    If PyLen = -1 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyLen", "argument has no len()"
    End If
End Function

' Return an iterator for the given object.
' Single-arg form:  iter(iterable)       — IIterable first, then fallbacks
' Two-arg form:     iter(callable, sentinel) — ISupplier + sentinel
Public Function PyIter(ByRef value As Variant, Optional ByVal sentinel As Variant) As IIterator
    ' 0. Nothing → not iterable
    If IsObject(value) Then
        If value Is Nothing Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyIter", "'Nothing' object is not iterable"
            Exit Function
        End If
    End If
    
    If IsMissing(sentinel) Then
        ' --- Single-arg form ---
        ' Fast path: explicit IIterable protocol
        If IsObject(value) Then
            If TypeOf value Is IIterable Then
                Dim iterable As IIterable
                Set iterable = value
                Set PyIter = iterable.Iter()
                Exit Function
            End If
        End If

        ' Fallback: try arrays, ISequence, IEnumVARIANT, etc.
        Dim fallback As IIterator
        Set fallback = TryCreateFallbackIterator(value)
        If Not (fallback Is Nothing) Then
            Set PyIter = fallback
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyIter", "argument is not iterable"
        End If
    Else
        ' Two-arg form: callable + sentinel
        If IsObject(value) Then
            If TypeOf value Is ISupplier Then
                Dim supplier As ISupplier
                Set supplier = value
                Dim supplierIter As SupplierIterator
                Set supplierIter = New SupplierIterator
                supplierIter.Init supplier, sentinel
                Set PyIter = supplierIter
            Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyIter", "object does not implement ISupplier"
            End If
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyIter", "argument is not a callable (ISupplier required)"
        End If
    End If
End Function

' Returns the next item from an iterator, matching Python's next(it) / next(it, default).
' One-arg form raises StopIteration on exhaustion; two-arg form returns default.
' Equivalent Python:
'   next(it)           → StopIteration raised
'   next(it, default)  → default returned
Public Function PyNext(ByRef iter As IIterator, Optional ByVal default As Variant) As Variant
    Dim item As Variant
    item = iter.NextItem()
    If IsError(item) Then
        If IsMissing(default) Then
            Err.Raise StopIterationCode.StopIteration, "Builtin.PyNext", _
                "iterator exhausted — no more items"
        Else
            PyNext = default
        End If
    Else
        PyNext = item
    End If
End Function

' Safe iteration helper that avoids the prime-then-loop pattern.
' Calls iter.NextItem() and returns:
'   True  → item contains the next value (use it inside the loop)
'   False → iterator is exhausted (item is undefined)
' Typical usage:
'   Dim item As Variant
'   Do While Builtin.PyTryNext(iter, item)
'       ' process item
'   Loop
Public Function PyTryNext(ByRef iter As IIterator, ByRef item As Variant) As Boolean
    Dim nextVal As Variant
    nextVal = iter.NextItem()
    If IsError(nextVal) Then
        PyTryNext = False
    Else
        PyTryNext = True
        item = nextVal
    End If
End Function

' Returns the truth value of a value, matching Python's bool() truth testing.
' https://docs.python.org/3/library/stdtypes.html#truth-value-testing
' Dispatch: primitives > PyLen > IBoolean (__bool__) > ILength (__len__) > True
Public Function PyBool(Optional ByRef value As Variant) As Boolean
    ' 0. Nothing → False (Python: bool(None) → False)
    If IsObject(value) Then
        If value Is Nothing Then
            PyBool = False
            Exit Function
        End If
    End If
    
    ' 1. No argument (missing), Empty, or Null → False
    ' IsMissing throws on undimensioned arrays (NULL SAFEARRAY), so we guard it.
    ' VB6 has no short-circuit Or, so we must check IsMissing in isolation.
    If Not IsArray(value) Then
        If IsMissing(value) Then
            PyBool = False
            Exit Function
        End If
    End If
    
    ' Empty or Null → False (safe on any Variant type)
    If IsEmpty(value) Or IsNull(value) Then
        PyBool = False
        Exit Function
    End If

    ' 2. Boolean → value itself
    If VarType(value) = vbBoolean Then
        PyBool = value
        Exit Function
    End If

    ' 3. Numeric types → non-zero is truthy
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            PyBool = (value <> 0)
            Exit Function
        Case vbDate
            PyBool = (CDbl(value) <> 0)
            Exit Function
    End Select

    ' 4. Use PyLenImpl for anything length-detectable
    ' Catches strings, arrays, Collections, ILength, Dictionary, etc.
    ' Returns -1 on failure, no On Error overhead.
    Dim length As Long
    length = PyLenImpl(value)
    If length <> -1 Then
        PyBool = (length > 0)
        Exit Function
    End If

    ' If PyLenImpl failed but value IS an array → undimensioned/empty → False
    If IsArray(value) Then
        PyBool = False
        Exit Function
    End If

    ' 5. Object — protocol dispatch
    If IsObject(value) Then
        ' Nothing → False
        If value Is Nothing Then
            PyBool = False
            Exit Function
        End If

        ' __bool__ protocol (IBoolean)
        If TypeOf value Is IBoolean Then
            Dim boolObj As IBoolean
            Set boolObj = value
            PyBool = boolObj.Bool()
            Exit Function
        End If

        ' __len__ protocol (ILength)
        If TypeOf value Is ILength Then
            Dim lenObj As ILength
            Set lenObj = value
            PyBool = (lenObj.Length() > 0)
            Exit Function
        End If

        ' Default: objects are truthy
        PyBool = True
        Exit Function
    End If

    ' 7. Fallback — try CBool, default True
    On Error Resume Next
    PyBool = CBool(value)
    If Err.Number <> 0 Then PyBool = True
    On Error GoTo 0
End Function

' Tries to create an IIterator from non-IIterable types using fallback mechanisms.
' Returns Nothing if no fallback applies to the given value.
Private Function TryCreateFallbackIterator(ByVal value As Variant) As IIterator
    ' 1. Native VB6 arrays
    If IsArray(value) Then
        Dim arrIter As ArrayIterator
        Set arrIter = New ArrayIterator
        arrIter.Init value
        Set TryCreateFallbackIterator = arrIter
        Exit Function
    End If
    
    ' 1b. Strings → character iterator
    If VarType(value) = vbString Then
        Dim strIter As StringIterator
        Set strIter = New StringIterator
        strIter.Init value
        Set TryCreateFallbackIterator = strIter
        Exit Function
    End If
    
    ' 2. ISequence protocol (indexable sequence)
    If IsObject(value) Then
        If TypeOf value Is ISequence Then
            Dim seq As ISequence
            Set seq = value
            Dim seqIter As SequenceIterator
            Set seqIter = New SequenceIterator
            seqIter.Init seq
            Set TryCreateFallbackIterator = seqIter
            Exit Function
        End If

        ' 3. COM enumerable objects (Collection, Dictionary, etc.)
        '    via IEnumVARIANT (For Each / _NewEnum)
        Dim enumOut As IUnknown
        Dim dispObj As Object: Set dispObj = value
        Dim enumHr As Long: enumHr = DispCall.TryGetEnum(dispObj, enumOut)
        If enumHr = 0 And Not (enumOut Is Nothing) Then
            Dim enumIter As EnumVARIANTIterator: Set enumIter = New EnumVARIANTIterator
            enumIter.InitFromEnumerator enumOut
            Set TryCreateFallbackIterator = enumIter
            Exit Function
        End If
    End If

    ' No fallback available
    Set TryCreateFallbackIterator = Nothing
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
Private Function PyLenByTypeOf(ByVal value As Variant, ByRef matched As Boolean) As Long
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

Private Function IsEnumerable(ByVal obj As Object) As Boolean
    Dim dispObj As Object: Set dispObj = obj
    Dim enumOut As IUnknown
    Dim callHr As Long: callHr = DispCall.TryGetEnum(dispObj, enumOut)
    If Err.Number = 0 Then
        IsEnumerable = True
    Else
        IsEnumerable = False
    End If
    Set enumOut = Nothing
End Function