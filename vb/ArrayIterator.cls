VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterable
Implements IIterator

Private m_data As Variant       ' array data (assigned in Init — VB6 copies by value)
Private m_lBound As Long        ' lower bound from Init (stable)
Private m_index As Long         ' current position, pre-incremented
Private m_exhausted As Boolean  ' True after last element returned

' Initialises the iterator with a copy of the array.
' Walking from LBound to UBound inclusive; bound is read fresh each NextItem call
' so the iterator adapts if the array storage changes.
Public Sub Init(ByRef arr As Variant)
    Ensure.IsTrue IsArray(arr), CommonHResults.TypeMismatch, _
        "ArrayIterator.Init", "arr must be an array"

    m_lBound = LBound(arr)
    m_index = m_lBound - 1
    m_data = arr
    m_exhausted = (m_lBound > UBound(arr))
End Sub

Private Function IIterator_NextItem() As Variant
    If m_exhausted Then
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
        Exit Function
    End If

    m_index = m_index + 1
    IIterator_NextItem = m_data(m_index)

    If m_index = UBound(m_data) Then
        m_exhausted = True
    End If
End Function

Private Function IIterator_Iter() As IIterator
    Set IIterator_Iter = Me
End Function

Private Function IIterable_Iter() As IIterator
    Set IIterable_Iter = Me
End Function