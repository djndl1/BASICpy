VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterable
Implements IIterator

Private m_source As String       ' source string
Private m_length As Long         ' Len(source) — precomputed, stable
Private m_index As Long          ' current position (0 = before first char)
Private m_exhausted As Boolean   ' True after last character returned

' Initialises the iterator with a string.
Public Sub Init(ByVal source As String)
    m_source = source
    m_length = Len(m_source)
    m_index = 0
    m_exhausted = (m_length = 0)
End Sub

Private Function IIterator_NextItem() As Variant
    If m_exhausted Then
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
        Exit Function
    End If
    
    m_index = m_index + 1
    IIterator_NextItem = Mid$(m_source, m_index, 1)
    
    If m_index = m_length Then
        m_exhausted = True
    End If
End Function

Private Function IIterator_Iter() As IIterator
    Set IIterator_Iter = Me
End Function

Private Function IIterable_Iter() As IIterator
    Set IIterable_Iter = Me
End Function
