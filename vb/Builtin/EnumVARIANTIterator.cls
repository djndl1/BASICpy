VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumVARIANTIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterable
Implements IIterator

' Wraps IEnumVARIANT via IEnumVARIANTWrapper for clean Next/Skip/Reset access.
Private m_wrapper As IEnumVARIANTWrapper
Private m_exhausted As Boolean       ' True once exhausted

' Backward-compatible initialiser — accepts any _NewEnum-supporting source.
Public Sub Init(ByRef source As Object)
    InitFromEnumerable source
End Sub

' Initialises from an already-obtained IEnumVARIANT.
Public Sub InitFromEnumerator(ByRef enumObj As IUnknown)
    m_exhausted = False
    Set m_wrapper = New IEnumVARIANTWrapper
    m_wrapper.InitFromEnumerator enumObj
End Sub

' Initialises from any _NewEnum-supporting source (calls TryGetEnum internally).
Public Sub InitFromEnumerable(ByRef source As Object)
    m_exhausted = False
    Set m_wrapper = New IEnumVARIANTWrapper
    m_wrapper.InitFromEnumerable source
End Sub

Private Sub Class_Terminate()
    Set m_wrapper = Nothing
End Sub

Private Function IIterator_NextItem() As Variant
    If m_exhausted Then
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
        Exit Function
    End If

    Builtin.LetSet IIterator_NextItem, m_wrapper.NextItem()

    If IsEmpty(IIterator_NextItem) Then
        m_exhausted = True
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
    End If
End Function

Private Function IIterator_Iter() As IIterator
    Set IIterator_Iter = Me
End Function

Private Function IIterable_Iter() As IIterator
    Set IIterable_Iter = Me
End Function