Attribute VB_Name = "AllAnyImpl"
Option Explicit

' Returns True if all elements of the iterable are truthy (or pass the predicate).
Public Function PyAllCore(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If

    Dim iter As IIterator: Set iter = Builtin.PyIter(value)
    PyAllCore = PyAllFromIter(iter, predicate)
End Function

' Returns True if any element of the iterable is truthy (or passes the predicate).
Public Function PyAnyCore(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If

    Dim iter As IIterator: Set iter = Builtin.PyIter(value)
    PyAnyCore = PyAnyFromIter(iter, predicate)
End Function

' Iterates through an IIterator, checking all elements pass the predicate.
Private Function PyAllFromIter(ByVal iterator As IIterator, ByVal predicate As IPredicate) As Boolean
    Dim item As Variant: Do While Builtin.PyTryNext(iterator, item)
        If Not predicate.Test(item) Then
            PyAllFromIter = False
            Exit Function
        End If
    Loop
    PyAllFromIter = True
End Function

' Iterates through an IIterator, checking if any element passes the predicate.
Private Function PyAnyFromIter(ByVal iterator As IIterator, ByVal predicate As IPredicate) As Boolean
    Dim item As Variant: Do While Builtin.PyTryNext(iterator, item)
        If predicate.Test(item) Then
            PyAnyFromIter = True
            Exit Function
        End If
    Loop
    PyAnyFromIter = False
End Function