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
    Dim item As Variant: Do While PyTryNext(iterator, item)
        If Not predicate.Test(item) Then
            PyAllFromIter = False
            Exit Function
        End If
    Loop
    PyAllFromIter = True
End Function

' Iterates through an IIterator, checking if any element passes the predicate.
Private Function PyAnyFromIter(ByVal iterator As IIterator, ByVal predicate As IPredicate) As Boolean
    Dim item As Variant: Do While PyTryNext(iterator, item)
        If predicate.Test(item) Then
            PyAnyFromIter = True
            Exit Function
        End If
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

' Converts a Variant to Long for PyRange, accepting only integers (Byte/Integer/Long)
' or objects implementing IIndex (__index__ equivalent). Delegates to PyInt for objects.
' Rejects floats and strings — range() requires exact integers.
Private Function AsIntegerIndex(ByVal value As Variant) As Long
    If IsObject(value) Then
        ' Use PyInt for IIndex protocol — recursion handled internally
        AsIntegerIndex = PyInt(value)
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

' Returns a RangeIterator matching Python's range().
' One-arg:  PyRange(stop)       → range(0, stop, 1)
' Two-arg:  PyRange(start, stop) → range(start, stop, 1)
' Three-arg: PyRange(start, stop, step) → range(start, stop, step)
Public Function PyRange(Optional ByVal arg1 As Variant, _
                        Optional ByVal arg2 As Variant, _
                        Optional ByVal arg3 As Long = 1) As RangeIterator
    Dim startVal As Long
    Dim stopVal As Long
    Dim stepVal As Long: stepVal = arg3
    
    If IsMissing(arg1) Then
        ' No args → empty range
        startVal = 0: stopVal = 0: stepVal = 1
    ElseIf IsMissing(arg2) Then
        ' Single-arg: PyRange(stop) → range(0, stop)
        startVal = 0
        stopVal = AsIntegerIndex(arg1)
        stepVal = 1
    Else
        ' Two/three-arg: PyRange(start, stop, step)
        startVal = AsIntegerIndex(arg1)
        stopVal = AsIntegerIndex(arg2)
        ' stepVal already assigned from arg3
    End If
    
    Dim r As New RangeIterator
    r.Init startVal, stopVal, stepVal
    Set PyRange = r
End Function

' ─── PyTuple ───

' Returns a new tuple (immutable sequence) wrapping the given value.
'   PyTuple()             → ()
'   PyTuple(arr)          → (arr(0), arr(1), ...) — copy of array
'   PyTuple(tupleObj)     → tupleObj (returned unchanged)
'   PyTuple("str")        → ('s', 't', 'r') — iterate chars
'   PyTuple(iterable)     → collect all items, wrap as tuple
Public Function PyTuple(Optional ByRef value As Variant) As Variant
    Dim tup As Tuple
    Dim objVal As Object
    
    ' No argument → empty tuple
    If IsMissing(value) Then
        Set tup = New Tuple
        tup.Init Array()
        Set PyTuple = tup
        Exit Function
    End If
    
    ' Already a Tuple → return as-is (same object)
    If IsObject(value) Then
        Set objVal = value
        If TypeOf objVal Is Tuple Then
            Set PyTuple = objVal
            Exit Function
        End If
    End If
    
    ' Array → copy into new Tuple
    If IsArray(value) Then
        Set tup = New Tuple
        tup.Init value
        Set PyTuple = tup
        Exit Function
    End If
    
    ' String → iterate characters
    If VarType(value) = vbString Then
        Dim str As String
        str = value
        Dim strLen As Long: strLen = Len(str)
        Dim strArr() As Variant
        Dim idx As Long
        If strLen > 0 Then
            ReDim strArr(0 To strLen - 1)
            For idx = 0 To strLen - 1
                strArr(idx) = Mid$(str, idx + 1, 1)
            Next idx
        Else
            strArr = Array()
        End If
        Set tup = New Tuple
        tup.Init strArr
        Set PyTuple = tup
        Exit Function
    End If
    
    ' Iterable (IIterable / IIterator / COM enumeration) → collect all items
    Dim iter As IIterator
    Set iter = PyIter(value)
    
    ' Collect into a growable array using Collection as a dynamic buffer
    Dim buffer As New Collection
    Dim item As Variant
    Do While PyTryNext(iter, item)
        buffer.Add item
    Loop
    
    ' Convert collection to array
    Dim bufLen As Long: bufLen = buffer.count
    Dim resultArr() As Variant
    Dim pos As Long
    If bufLen > 0 Then
        ReDim resultArr(0 To bufLen - 1)
        For pos = 1 To bufLen
            resultArr(pos - 1) = buffer(pos)
        Next pos
    Else
        resultArr = Array()
    End If
    
    Set tup = New Tuple
    tup.Init resultArr
    Set PyTuple = tup
End Function

' ─── VBArray ───

' Returns a new VBArray (mutable sequence) wrapping the given value.
'   VBArray()             → []
'   VBArray(arr)          → [arr(0), arr(1), ...] — copy of array
'   VBArray(str)          → ['s', 't', 'r'] — iterate chars
'   VBArray(iterable)     → collect all items
Public Function VBArrayBuiltin(Optional ByRef value As Variant) As VBArray
    Dim vba As VBArray
    Dim objVal As Object
    
    ' No argument → empty
    If IsMissing(value) Then
        Set VBArrayBuiltin = New VBArray
        Exit Function
    End If
    
    ' Already a VBArray → use Copy (single array copy vs O(n) Append)
    If IsObject(value) Then
        Set objVal = value
        If TypeOf objVal Is VBArray Then
            Dim src As VBArray
            Set src = objVal
            Set VBArrayBuiltin = src.Copy()
            Exit Function
        End If
    End If
    
    ' Array → copy into new VBArray
    If IsArray(value) Then
        Set vba = New VBArray
        vba.Init value
        Set VBArrayBuiltin = vba
        Exit Function
    End If
    
    ' String → iterate characters
    If VarType(value) = vbString Then
        Set vba = New VBArray
        Dim str As String: str = value
        Dim strLen As Long: strLen = Len(str)
        Dim idx As Long
        For idx = 1 To strLen
            vba.Append Mid$(str, idx, 1)
        Next idx
        Set VBArrayBuiltin = vba
        Exit Function
    End If
    
    ' Iterable → collect all items
    Dim iter As IIterator: Set iter = PyIter(value)
    Set vba = New VBArray
    Dim item As Variant: Do While PyTryNext(iter, item)
        vba.Append item
    Loop
    Set VBArrayBuiltin = vba
End Function

' ─── PyInt helpers ───

' Converts a single character to its digit value (0-35).
' Returns -1 for non-digit characters.
Private Function CharToDigit(ByVal c As String) As Long
    Dim code As Long
    code = Asc(c)
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
Private Function ParseIntString(ByVal s As String, ByVal base As Long) As Long
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

' Returns an integer matching Python's int().
'   PyInt()              → 0
'   PyInt(number)        → truncates towards zero (Fix for floats)
'   PyInt(string)        → parses in base 10
'   PyInt(string, base)  → parses in given base (0 or 2-36)
'   PyInt(obj)           → calls IIndex protocol
Public Function PyInt(Optional ByVal value As Variant, Optional ByVal base As Variant) As Variant
    If IsArray(value) Then
        ' Reject arrays before IsMissing check (IsMissing throws on undimensioned arrays)
        Ensure.IsTrue False, CommonHResults.TypeMismatch, _
            "Builtin.PyInt", "argument must be a number or string, not an array"
        
    ElseIf IsMissing(value) Then
        ' int() → 0
        PyInt = 0
        
    ElseIf Not IsMissing(base) Then
        ' int(string, base) — force string path
        PyInt = ParseIntString(CStr(value), AsIntegerIndex(base))

    ElseIf IsObject(value) Then
        ' Object → IIndex protocol
        If TypeOf value Is IIndex Then
            Dim idx As IIndex
            Set idx = value
            PyInt = PyInt(idx.index())        ' recurse with unwrapped value
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyInt", "object does not implement IIndex"
        End If

    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyInt = value

            Case vbSingle, vbDouble, vbCurrency, vbDecimal
                PyInt = Fix(value)              ' truncate towards zero

            Case vbString
                PyInt = ParseIntString(value, 10)   ' default base 10

            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyInt", "argument must be a number or string"
        End Select
    End If
End Function

' Returns True if item is in container, matching Python's `in` operator.
' Fast path:  IContains protocol (__contains__).
' Type-name:  Known COM containers (Dictionary, Hashtable, etc.) via direct methods.
' Dynamic:    Late-bound Contains/ContainsKey/Exists for unknown types.
' Fallback:   Iteration via PyIter.
Public Function PyIn(ByRef item As Variant, ByRef container As Variant) As Boolean
    ' 0. Undimensioned array container → empty, nothing to find
    If IsArray(container) Then
        If PyLenImpl(container) = -1 Then
            PyIn = False
            Exit Function
        End If
    End If
    
    ' 1. IContains protocol (constant-time when available)
    If IsObject(container) Then
        If TypeOf container Is IContains Then
            Dim c As IContains
            Set c = container
            PyIn = c.Contains(item)
            Exit Function
        End If
    End If
    
    ' 1b. String container → substring matching (Python: y.find(x) != -1)
    ' TODO: also support bytes / bytearray (substring matching for byte sequences)
    If VarType(container) = vbString Then
        If VarType(item) <> vbString Then
            Ensure.IsTrue False, CommonHResults.TypeMismatch, _
                "Builtin.PyIn", "'in <string>' requires string as left operand"
        End If
        PyIn = (InStr(1, container, item, vbBinaryCompare) > 0)
        Exit Function
    End If
    
    ' 2. Type-name fast path — known COM containers with O(1) lookup
    If IsObject(container) Then
        Dim dispObj As Object: Set dispObj = container
        Dim matched As Boolean: matched = False
        PyIn = PyInByTypeName(item, dispObj, matched)
        If matched Then Exit Function
    End If
    
    ' 3. Iteration fallback
    Dim element As Variant: Do While PyTryNext(PyIter(container), element)
        If element = item Then
            PyIn = True
            Exit Function
        End If
    Loop
    
    ' 4. Dynamic contains check — last resort for unknown types
    '     that support Contains / ContainsKey / Exists
    If IsObject(container) Then
        Dim dynResult As Variant
        dynResult = PyInDynamicCheck(item, dispObj)
        If Not IsEmpty(dynResult) Then
            PyIn = dynResult
            Exit Function
        End If
    End If
    
    PyIn = False
End Function

' Matches known COM container types by TypeName and uses their O(1) lookup method.
' Returns True/False for known containers (sets matched=True).
' Returns False with matched=False for unknown types (caller falls through).
Private Function PyInByTypeName(ByRef item As Variant, ByRef container As Object, _
                                ByRef matched As Boolean) As Boolean
    matched = True
    Select Case TypeName(container)
        ' --- Scripting.Dictionary — .Exists(key) ---
        Case "Dictionary"
            PyInByTypeName = container.Exists(item)
        
        ' --- .NET System.Collections ---
        ' Hashtable uses .ContainsKey  (IDictionary)
        Case "Hashtable"
            PyInByTypeName = container.ContainsKey(item)
        
        ' SortedList implements IDictionary — .ContainsKey
        Case "SortedList"
            PyInByTypeName = container.ContainsKey(item)
        
        ' ICollection / IList types — .Contains(item)
        Case "ArrayList", "ICollection", "BitArray", "Stack", "Queue"
            PyInByTypeName = container.Contains(item)
        
        ' --- Known types without O(1) membership — signal matched=False ---
        Case "Collection"
            matched = False
        Case "ListBox", "ComboBox", "DirListBox", "FileListBox"
            matched = False
        Case "Form", "MDIForm", "PropertyPage"
            matched = False
        Case "Recordset"
            matched = False
        
        ' --- Unknown type — signal caller to try dynamic or iteration ---
        Case Else
            matched = False
    End Select
End Function

' Tries late-bound Contains, ContainsKey, and Exists on unknown object types.
' Returns Variant tri-state:
'   True / False  — authoritative answer (container supports one of these methods)
'   Empty         — none of the methods exist, caller should use iteration
Private Function PyInDynamicCheck(ByRef item As Variant, ByRef container As Object) As Variant
    On Error Resume Next
    
    PyInDynamicCheck = container.Contains(item)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    
    PyInDynamicCheck = container.ContainsKey(item)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    
    PyInDynamicCheck = container.Exists(item)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    
    PyInDynamicCheck = Empty
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
                Set PyIter = iterable.iter()
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
    Dim Length As Long
    Length = PyLenImpl(value)
    If Length <> -1 Then
        PyBool = (Length > 0)
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
