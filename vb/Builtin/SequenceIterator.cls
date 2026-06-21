VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SequenceIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterable
Implements IIterator

Private m_sequence As ISequence   ' source sequence (held by reference)
Private m_index As Long           ' current position, pre-incremented
Private m_exhausted As Boolean    ' True after last element returned

' Initialises the iterator with an ISequence.
' Walks from index 0 to Length()-1 inclusive.
Public Sub Init(ByRef sequence As ISequence)
    Set m_sequence = sequence
    m_index = -1
    m_exhausted = (m_sequence.Length() <= 0)
End Sub

Private Function IIterator_NextItem() As Variant
    If m_exhausted Then
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
        Exit Function
    End If
    
    m_index = m_index + 1
    
    If m_index >= m_sequence.Length() Then
        m_exhausted = True
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
    Else
        Dim raw As Variant
        Builtin.LetSet raw, m_sequence.GetItem(m_index)
        Builtin.LetSet IIterator_NextItem, raw
    End If
End Function

Private Function IIterator_Iter() As IIterator
    Set IIterator_Iter = Me
End Function

Private Function IIterable_Iter() As IIterator
    Set IIterable_Iter = Me
End Function
