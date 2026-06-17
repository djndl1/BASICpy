VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterable

' Returns the next element as a Variant.
' Returns CVErr(CBasicPyTypes.StopIterationCode.StopIteration) when exhausted.
' Once exhausted, every subsequent call must return the same sentinel.
Public Function NextItem() As Variant
End Function

' Returns self — an iterator is iterable.
' Declared directly on IIterator so Iter() is available without casting to IIterable.
Public Function Iter() As IIterator
End Function

' Implements IIterable.Iter
Private Function IIterable_Iter() As IIterator
    Set IIterable_Iter = Me
End Function
