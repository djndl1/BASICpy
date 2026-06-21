Attribute VB_Name = "NextImpl"
Option Explicit

' Returns the next item from an iterator, matching Python's next(it) / next(it, default).
' One-arg form raises StopIteration on exhaustion; two-arg form returns default.
' Equivalent Python:
'   next(it)           → StopIteration raised
'   next(it, default)  → default returned
Public Function PyNextCore(ByRef iter As IIterator, Optional ByVal default As Variant) As Variant
    Dim item As Variant
    Dim raw As Variant: Builtin.LetSet raw, iter.NextItem()
    Builtin.LetSet item, raw
    If IsError(item) Then
        If IsMissing(default) Then
            Err.Raise StopIterationCode.StopIteration, "Builtin.PyNext", _
                "iterator exhausted — no more items"
        Else
            Builtin.LetSet PyNextCore, default
        End If
    Else
        Builtin.LetSet PyNextCore, item
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
Public Function PyTryNextCore(ByRef iter As IIterator, ByRef item As Variant) As Boolean
    Dim nextVal As Variant
    Dim raw As Variant: Builtin.LetSet raw, iter.NextItem()
    Builtin.LetSet nextVal, raw
    If IsError(nextVal) Then
        PyTryNextCore = False
    Else
        PyTryNextCore = True
        Builtin.LetSet item, nextVal
    End If
End Function
