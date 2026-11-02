Attribute VB_Name = "MaxMinImpl"
Option Explicit

' Core implementation for PyMax — returns the largest item in an iterable,
' matching Python's max().
'
' Python semantics:
'   max(1, 2, 3)              → 3
'   max([1, 2, 3])            → 3
'   max("hello")              → 'o'
'   max([])                   → raises ValueError
'
' VB6: Single iterable argument. Multi-arg form (max(a, b, c)) not supported
'      due to VB6's lack of variadic parameters — wrap args in an array instead.
'
' Optional key function: if provided, elements are compared by their key result.
' This matches Python's key parameter (an ordering function like list.sort's key),
' NOT a boolean predicate.
Public Function PyMaxCore(ByRef value As Variant, Optional ByVal key As IFunction) As Variant
    PyMaxCore = PyMaxMinCore("Builtin.PyMax", "max() arg is an empty sequence", value, True, key)
End Function

' Core implementation for PyMin — returns the smallest item in an iterable,
' matching Python's min().
'
' Python semantics:
'   min(1, 2, 3)              → 1
'   min([1, 2, 3])            → 1
'   min("hello")              → 'e'
'   min([])                   → raises ValueError
'
' VB6: Single iterable argument.
Public Function PyMinCore(ByRef value As Variant, Optional ByVal key As IFunction) As Variant
    PyMinCore = PyMaxMinCore("Builtin.PyMin", "min() arg is an empty sequence", value, False, key)
End Function

' Shared core implementation for PyMax/PyMin.
'
' Algorithm: iterate once, track current best item. For each subsequent item:
'   findMax = True  → keep item if item > currentBest  (Python __gt__)
'   findMax = False → keep item if item < currentBest  (Python __lt__)
'
' When key is provided, compare key.Apply(item) vs key.Apply(currentBest)
' instead of the raw items. This matches Python's key parameter for max/min.
Private Function PyMaxMinCore(ByVal source As String, ByVal emptyMsg As String, _
                              ByRef value As Variant, ByVal findMax As Boolean, _
                              Optional ByVal key As IFunction) As Variant
    Dim iter As IIterator: Set iter = Builtin.PyIter(value)

    Dim item As Variant: Dim hasItem As Boolean: hasItem = Builtin.PyTryNext(iter, item)

    Ensure.IsTrue hasItem, CommonHResults.InvalidProcedureCall, source, emptyMsg

    Dim currentBest As Variant: Builtin.LetSet currentBest, item

    Dim currentBestKey As Variant
    If Not (key Is Nothing) Then
        Builtin.LetSet currentBestKey, key.Apply(item)
    End If

    Do While Builtin.PyTryNext(iter, item)
        Dim itemKey As Variant
        If Not (key Is Nothing) Then
            Builtin.LetSet itemKey, key.Apply(item)
        End If

        Dim keep As Boolean
        If Not (key Is Nothing) Then
            keep = IsOrdered(itemKey, currentBestKey, findMax)
        Else
            keep = IsOrdered(item, currentBest, findMax)
        End If

        If keep Then
            Builtin.LetSet currentBest, item
            If Not (key Is Nothing) Then
                Builtin.LetSet currentBestKey, itemKey
            End If
        End If
    Loop

    Builtin.LetSet PyMaxMinCore, currentBest
End Function

' Python rich comparison (__lt__ / __gt__). Returns True if:
'   findMax = False → a < b  (used by min)
'   findMax = True  → a > b  (used by max)
'
' Resolution order:
'   1. If a implements ILessThan/IGreaterThan → a.LessThan(b) / a.GreaterThan(b)
'   2. Else if b implements ILessThan/IGreaterThan → Not b.LessThan(a) / Not b.GreaterThan(a)
'   3. Else type-based comparison (numeric, string; None/mixed → TypeError)
Private Function IsOrdered(ByVal a As Variant, ByVal b As Variant, _
                           ByVal findMax As Boolean) As Boolean
    Dim aObj As Object
    Dim bObj As Object
    If IsObject(a) Then Set aObj = a
    If IsObject(b) Then Set bObj = b

    If findMax Then
        ' Python max(): a > b
        If Not (aObj Is Nothing) And TypeOf aObj Is IGreaterThan Then
            IsOrdered = aObj.GreaterThan(b)
            Exit Function
        End If
        If Not (bObj Is Nothing) And TypeOf bObj Is IGreaterThan Then
            IsOrdered = Not bObj.GreaterThan(a)
            Exit Function
        End If
    Else
        ' Python min(): a < b
        If Not (aObj Is Nothing) And TypeOf aObj Is ILessThan Then
            IsOrdered = aObj.LessThan(b)
            Exit Function
        End If
        If Not (bObj Is Nothing) And TypeOf bObj Is ILessThan Then
            IsOrdered = Not bObj.LessThan(a)
            Exit Function
        End If
    End If

    IsOrdered = CompareByType(a, b, findMax)
End Function

' Type-based comparison. Returns True if:
'   findMax = False → a < b
'   findMax = True  → a > b
'
'   - Nothing/Empty (Python None) → TypeError (not comparable)
'   - Numeric vs numeric → numeric comparison (Decimal vs Decimal exact)
'   - String vs string → lexicographic
'   - Mixed types → TypeError (Python 3 semantics)
Private Function CompareByType(ByVal a As Variant, ByVal b As Variant, _
                               ByVal findMax As Boolean) As Boolean
    Dim aIsNone As Boolean
    Dim bIsNone As Boolean
    aIsNone = IsNothingLike(a)
    bIsNone = IsNothingLike(b)

    If aIsNone Or bIsNone Then
        If findMax Then
            Err.Raise vbObjectError + 1001, "CompareByType", _
                "TypeError: '>' not supported between instances of 'NoneType' and 'NoneType'"
        Else
            Err.Raise vbObjectError + 1001, "CompareByType", _
                "TypeError: '<' not supported between instances of 'NoneType' and 'NoneType'"
        End If
        Exit Function
    End If

    Dim aType As VbVarType
    Dim bType As VbVarType
    aType = VarType(a)
    bType = VarType(b)

    ' Both numeric → numeric comparison
    If IsNumericType(aType) And IsNumericType(bType) Then
        Dim result As Boolean
        If aType = vbDecimal And bType = vbDecimal Then
            If findMax Then
                result = (a > b)
            Else
                result = (a < b)
            End If
        Else
            If findMax Then
                result = (CDbl(a) > CDbl(b))
            Else
                result = (CDbl(a) < CDbl(b))
            End If
        End If
        CompareByType = result
        Exit Function
    End If

    ' Both strings → lexicographic
    If aType = vbString And bType = vbString Then
        If findMax Then
            CompareByType = (a > b)
        Else
            CompareByType = (a < b)
        End If
        Exit Function
    End If

    ' Mixed types → TypeError
    Dim op As String
    If findMax Then op = ">" Else op = "<"
    Err.Raise vbObjectError + 1002, "CompareByType", _
        "TypeError: '" & op & "' not supported between instances of '" & _
        VarTypeName(aType) & "' and '" & VarTypeName(bType) & "'"
End Function

' Returns True if the variant holds Nothing or Empty (Python None equivalent).
Private Function IsNothingLike(ByVal v As Variant) As Boolean
    If IsObject(v) Then
        If v Is Nothing Then
            IsNothingLike = True
            Exit Function
        End If
    End If
    If VarType(v) = vbEmpty Then
        IsNothingLike = True
        Exit Function
    End If
    IsNothingLike = False
End Function

' Returns True if the VarType is a numeric type.
Private Function IsNumericType(ByVal vt As VbVarType) As Boolean
    Select Case vt
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsNumericType = True
        Case Else
            IsNumericType = False
    End Select
End Function

' Returns a readable name for a VarType constant.
Private Function VarTypeName(ByVal vt As VbVarType) As String
    Select Case vt
        Case vbEmpty:      VarTypeName = "Empty"
        Case vbNull:       VarTypeName = "Null"
        Case vbInteger:    VarTypeName = "Integer"
        Case vbLong:       VarTypeName = "Long"
        Case vbSingle:     VarTypeName = "Single"
        Case vbDouble:     VarTypeName = "Double"
        Case vbCurrency:   VarTypeName = "Currency"
        Case vbDecimal:    VarTypeName = "Decimal"
        Case vbDate:       VarTypeName = "Date"
        Case vbString:     VarTypeName = "String"
        Case vbObject:     VarTypeName = "Object"
        Case vbError:      VarTypeName = "Error"
        Case vbBoolean:    VarTypeName = "Boolean"
        Case vbVariant:    VarTypeName = "Variant"
        Case vbDataObject: VarTypeName = "DataObject"
        Case vbByte:       VarTypeName = "Byte"
        Case VBArray:      VarTypeName = "Array"
        Case Else:         VarTypeName = "Unknown(" & CStr(vt) & ")"
    End Select
End Function