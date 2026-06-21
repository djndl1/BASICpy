VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ILength

' Returns the element at the given index (0-based, matching Python __getitem__).
Public Function GetItem(ByVal Index As Long) As Variant
End Function

' ISequence exposes its own length.
Public Function Length() As Long
End Function

' Returns the number of occurrences of value in the sequence (sequence.count).
Public Function count(ByVal value As Variant) As Long
End Function

' Returns the index of the first occurrence of value (sequence.index).
' Optional start/stop bounds for efficient sub-range search.
' Raises if value is not found.
Public Function Index(ByVal value As Variant, Optional ByVal start As Variant, Optional ByVal stopAt As Variant) As Long
End Function

' Delegates ILength.Length to ISequence's own Length.
Private Function ILength_Length() As Long
    ILength_Length = Length()
End Function
