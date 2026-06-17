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
Public Function GetItem(ByVal index As Long) As Variant
End Function

' ISequence exposes its own length.
Public Function Length() As Long
End Function

' Delegates ILength.Length to ISequence's own Length.
Private Function ILength_Length() As Long
    ILength_Length = Length()
End Function
