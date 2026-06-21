Attribute VB_Name = "IterCoreImpl"
Option Explicit

' Core implementation for PyIter — returns an IIterator for the given object.
Public Function PyIterCore(ByRef value As Variant, Optional ByVal sentinel As Variant) As IIterator
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
        If IsObject(value) Then
            If TypeOf value Is IIterable Then
                Dim iterable As IIterable: Set iterable = value
                Set PyIterCore = iterable.iter()
                Exit Function
            End If
        End If
        
        ' Fallback: try arrays, ISequence, IEnumVARIANT, etc.
        Dim fallback As IIterator: Set fallback = TryCreateFallbackIterator(value)
        If Not (fallback Is Nothing) Then
            Set PyIterCore = fallback
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyIter", "argument is not iterable"
        End If
    Else
        ' Two-arg form: callable + sentinel
        If IsObject(value) Then
            If TypeOf value Is ISupplier Then
                Dim supplier As ISupplier: Set supplier = value
                Dim supplierIter As SupplierIterator: Set supplierIter = New SupplierIterator
                supplierIter.Init supplier, sentinel
                Set PyIterCore = supplierIter
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
