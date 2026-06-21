VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupplierIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterable
Implements IIterator

Private m_supplier As ISupplier
Private m_sentinel As Variant
Private m_exhausted As Boolean

Public Sub Init(ByVal supplier As ISupplier, ByVal sentinel As Variant)
    Set m_supplier = supplier
    m_sentinel = sentinel
    m_exhausted = False
End Sub

Private Function IIterator_NextItem() As Variant
    If m_exhausted Then
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
        Exit Function
    End If

    Dim result As Variant
    result = m_supplier.GetResult()

    If IsObject(result) Then
        If result Is m_sentinel Then
            m_exhausted = True
            IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
            Exit Function
        End If
    Else
        If result = m_sentinel Then
            m_exhausted = True
            IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
            Exit Function
        End If
    End If

    IIterator_NextItem = result
End Function

Private Function IIterator_Iter() As IIterator
    Set IIterator_Iter = Me
End Function

Private Function IIterable_Iter() As IIterator
    Set IIterable_Iter = Me
End Function
