Attribute VB_Name = "IterCoreImpl"
Option Explicit

' Core implementation for PyIter — returns an IIterator for the given object.
Public Function PyIterCore(ByRef value As Variant, Optional ByVal sentinel As Variant) As IIterator
    ' Guard: Nothing is not iterable
    If IsObject(value) Then
        If value Is Nothing Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyIter", "'Nothing' object is not iterable"
            Exit Function
        End If
    End If

    ' Guard: two-arg form (callable + sentinel)
    If Not IsMissing(sentinel) Then
        ' Guard: must be an object (callable)
        If Not IsObject(value) Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyIter", "argument is not a callable (ISupplier required)"
            Exit Function
        End If

        ' Guard: must implement ISupplier
        If Not TypeOf value Is ISupplier Then
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyIter", "object does not implement ISupplier"
            Exit Function
        End If

        ' Happy path: create supplier iterator
        Dim supplier As ISupplier: Set supplier = value
        Dim supplierIter As SupplierIterator: Set supplierIter = New SupplierIterator
        supplierIter.Init supplier, sentinel
        Set PyIterCore = supplierIter
        Exit Function
    End If

    ' --- Single-arg form ---
    ' Guard 1: IIterable fast path
    If IsObject(value) Then
        If TypeOf value Is IIterable Then
            Dim iterable As IIterable: Set iterable = value
            Set PyIterCore = iterable.iter()
            Exit Function
        End If
    End If

    ' Guard 2: Fallback strategies (arrays, strings, ISequence, COM enumerable)
    Dim fallback As IIterator: Set fallback = TryCreateFallbackIterator(value)
    If Not (fallback Is Nothing) Then
        Set PyIterCore = fallback
        Exit Function
    End If

    ' Error: not iterable
    Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
        "Builtin.PyIter", "argument is not iterable"
End Function

' Creates an IIterator for a value using fallback strategies:
' 1. Native VB6 arrays → ArrayIterator
' 2. Strings → StringIterator (character-by-character)
' 3. ISequence protocol → SequenceIterator
' 4. COM enumerable → EnumVARIANTIterator (via _NewEnum / IEnumVARIANT)
' Returns Nothing if no strategy applies.
Public Function TryCreateFallbackIterator(ByVal value As Variant) As IIterator
    ' 1. Native VB6 arrays
    If IsArray(value) Then
        Dim arrIter As ArrayIterator: Set arrIter = New ArrayIterator
        arrIter.Init value
        Set TryCreateFallbackIterator = arrIter
        Exit Function
    End If

    ' 1b. Strings — character iterator
    If VarType(value) = vbString Then
        Dim strIter As StringIterator: Set strIter = New StringIterator
        strIter.Init value
        Set TryCreateFallbackIterator = strIter
        Exit Function
    End If

    ' Guard: only objects can be ISequence or COM enumerable
    If Not IsObject(value) Then
        Set TryCreateFallbackIterator = Nothing
        Exit Function
    End If

    ' 2. ISequence protocol (indexable sequence)
    If TypeOf value Is ISequence Then
        Dim seq As ISequence: Set seq = value
        Dim seqIter As SequenceIterator: Set seqIter = New SequenceIterator
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

    ' No fallback available
    Set TryCreateFallbackIterator = Nothing
End Function

' Checks if an object supports COM enumeration (For Each / _NewEnum / IEnumVARIANT).
Public Function IsEnumerable(ByVal obj As Object) As Boolean
    Dim dispObj As Object: Set dispObj = obj
    Dim enumOut As IUnknown
    Dim callHr As Long: callHr = DispCall.TryGetEnum(dispObj, enumOut)

    ' Guard: error during enumeration check
    If Err.Number <> 0 Then
        IsEnumerable = False
        Set enumOut = Nothing
        Exit Function
    End If

    IsEnumerable = True
    Set enumOut = Nothing
End Function