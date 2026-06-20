VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMutableSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Appends value to the end of the sequence.
Public Sub Append(ByRef value As Variant)
End Sub

' Removes all items from the sequence.
Public Sub Clear()
End Sub

' Returns a shallow copy of the sequence (new mutable sequence with same items).
Public Function Copy() As Variant
End Function

' Extends the sequence with the contents of iterable.
Public Sub Extend(ByRef iterable As Variant)
End Sub

' Inserts value at the given index. Shifts subsequent elements to the right.
Public Sub Insert(ByVal index As Long, ByRef value As Variant)
End Sub

' Retrieves the item at index and removes it from the sequence.
' Default index is -1 (last item).
Public Function Pop(Optional ByVal index As Long = -1) As Variant
End Function

' Removes the first occurrence of value from the sequence.
' Raises if value is not found.
Public Sub Remove(ByRef value As Variant)
End Sub

' Reverses the items of the sequence in place.
Public Sub Reverse()
End Sub
