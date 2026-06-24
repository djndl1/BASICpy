Attribute VB_Name = "InCoreImpl"
Option Explicit

' Core implementation for PyIn — returns True if item is in container.
Public Function PyInCore(ByRef item As Variant, ByRef container As Variant) As Boolean
    ' 0. Undimensioned array container → empty, nothing to find
    If IsArray(container) Then
        If PyLenImpl(container) = -1 Then
            PyInCore = False
            Exit Function
        End If
    End If

    ' 1. IContains protocol (constant-time when available)
    If IsObject(container) Then
        If TypeOf container Is IContains Then
            Dim c As IContains: Set c = container
            PyInCore = c.Contains(item)
            Exit Function
        End If
    End If

    ' 1b. String container → substring matching (Python: y.find(x) != -1)
    If VarType(container) = vbString Then
        If VarType(item) <> vbString Then
            Ensure.IsTrue False, CommonHResults.TypeMismatch, _
                "Builtin.PyIn", "'in <string>' requires string as left operand"
        End If
        PyInCore = (InStr(1, container, item, vbBinaryCompare) > 0)
        Exit Function
    End If

    ' 2. Type-name fast path — known COM containers with O(1) lookup
    If IsObject(container) Then
        Dim dispObj As Object: Set dispObj = container
        Dim matched As Boolean: matched = False
        PyInCore = PyInByTypeName(item, dispObj, matched)
        If matched Then Exit Function
    End If

    ' 3. Iteration fallback
    Dim iter As IIterator: Set iter = Builtin.PyIter(container)
    Dim element As Variant: Do While Builtin.PyTryNext(iter, element)
        If element = item Then
            PyInCore = True
            Exit Function
        End If
    Loop

    ' 4. Dynamic contains check — last resort for unknown types
    If IsObject(container) Then
        Dim dynResult As Variant: dynResult = PyInDynamicCheck(item, dispObj)
        If Not IsEmpty(dynResult) Then
            PyInCore = dynResult
            Exit Function
        End If
    End If

    PyInCore = False
End Function

' Fast-path membership checks for known COM container types by TypeName.
' Sets matched=True and returns True/False if type is known with O(1) membership.
' Sets matched=False for known types without O(1) membership (Collection, ListBox, etc.)
' or for unknown types (caller falls back to iteration or dynamic check).
Public Function PyInByTypeName(ByRef item As Variant, ByRef container As Object, _
                                ByRef matched As Boolean) As Boolean
    matched = True
    Select Case TypeName(container)
        ' --- Scripting.Dictionary — .Exists(key) ---
        Case "Dictionary"
            PyInByTypeName = container.Exists(item)

        ' --- .NET System.Collections ---
        Case "Hashtable", "SortedList"
            PyInByTypeName = container.ContainsKey(item)

        ' ICollection / IList types — .Contains(item)
        Case "ArrayList", "ICollection", "BitArray", "Stack", "Queue"
            PyInByTypeName = container.Contains(item)
        ' --- Unknown type — signal caller to try dynamic or iteration ---
        Case Else
            matched = False
    End Select
End Function

' Tries late-bound Contains, ContainsKey, and Exists on unknown object types.
' Returns Variant tri-state:
'   True / False  — authoritative answer (container supports one of these methods)
'   Empty         — none of the methods exist, caller should use iteration
Public Function PyInDynamicCheck(ByRef item As Variant, ByRef container As Object) As Variant
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