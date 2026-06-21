Attribute VB_Name = "IterImpl"
Option Explicit

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
    
    ' 2. ISequence protocol (indexable sequence)
    If IsObject(value) Then
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
    End If
    
    ' No fallback available
    Set TryCreateFallbackIterator = Nothing
End Function

' Checks if an object supports COM enumeration (For Each / _NewEnum / IEnumVARIANT).
Public Function IsEnumerable(ByVal obj As Object) As Boolean
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
