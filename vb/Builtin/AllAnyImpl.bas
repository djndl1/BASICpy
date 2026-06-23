Attribute VB_Name = "AllAnyImpl"
Option Explicit

' Returns True if all elements of the iterable are truthy (or pass the predicate).
Public Function PyAllCore(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If
    Dim iter As IIterator: Set iter = Builtin.PyIter(value)
    Dim item As Variant: Do While Builtin.PyTryNext(iter, item)
        If Not predicate.Test(item) Then
            PyAllCore = False
            Exit Function
        End If
    Loop
    PyAllCore = True
End Function

' Returns True if any element of the iterable is truthy (or passes the predicate).
Public Function PyAnyCore(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Set predicate = IsTruthy
    End If

    Dim iter As IIterator: Set iter = Builtin.PyIter(value)
    Dim item As Variant: Do While Builtin.PyTryNext(iter, item)
        If predicate.Test(item) Then
            PyAnyCore = True
            Exit Function
        End If
    Loop
    PyAnyCore = False
End Function
