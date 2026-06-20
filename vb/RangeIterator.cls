VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangeIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IContains
Implements ISequence
Implements ILength
Implements IIterable

Private m_start As Long
Private m_stop As Long
Private m_step As Long
Private m_length As Long

' Initialises the range with start, stop, step.
' Step must not be zero (raises error).
Public Sub Init(ByVal RangeStart As Long, _
                ByVal RangeStop As Long, _
                Optional ByVal RangeStep As Long = 1)
    If RangeStep = 0 Then
        Err.Raise 5, "RangeIterator.Init", "step must not be zero"
    End If

    m_start = RangeStart
    m_stop = RangeStop
    m_step = RangeStep

    ' Precompute length using Python's range length formula
    If m_step > 0 Then
        If m_start >= m_stop Then
            m_length = 0
        Else
            m_length = 1 + (m_stop - 1 - m_start) \ m_step
        End If
    Else
        If m_start <= m_stop Then
            m_length = 0
        Else
            m_length = 1 + (m_stop + 1 - m_start) \ m_step
        End If
    End If
End Sub

' Read-only properties matching Python range.start, range.stop, range.step
Public Property Get RangeStart() As Long
    RangeStart = m_start
End Property

Public Property Get RangeStop() As Long
    RangeStop = m_stop
End Property

Public Property Get RangeStep() As Long
    RangeStep = m_step
End Property

' Constant-time containment: returns True if value is in the range.
Public Function Contains(ByVal value As Long) As Boolean
    If m_length = 0 Then
        Contains = False
        Exit Function
    End If

    Dim diff As Long
    diff = value - m_start

    ' Must be evenly divisible by step
    If diff Mod m_step <> 0 Then
        Contains = False
        Exit Function
    End If

    ' Index must be within [0, length)
    Dim Index As Long
    Index = diff \ m_step

    Contains = (Index >= 0 And Index < m_length)
End Function

' Returns the index of value in the range, or raises if not present.
Public Function Index(ByVal value As Long) As Long
    If Not Contains(value) Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "RangeIterator.Index", "value not in range"
    End If
    Index = (value - m_start) \ m_step
End Function

' Default Item property — enables rangeObj(index) syntax (VB6 default member).
Public Property Get item(ByVal Index As Long) As Variant
Attribute item.VB_UserMemId = 0
    item = ISequence_GetItem(Index)
End Property

' ─── IContains — delegates to the public Contains(Long) ───

Private Function IContains_Contains(ByVal value As Variant) As Boolean
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong
            IContains_Contains = Contains(CLng(value))
        Case Else
            IContains_Contains = False
    End Select
End Function

' ─── ISequence ───

Private Function ISequence_GetItem(ByVal Index As Long) As Variant
    Dim actualIndex As Long
    actualIndex = Index
    If actualIndex < 0 Then
        actualIndex = m_length + actualIndex
    End If

    If actualIndex < 0 Or actualIndex >= m_length Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "RangeIterator.GetItem", "index out of range"
    End If

    ISequence_GetItem = m_start + m_step * actualIndex
End Function

Private Function ISequence_Length() As Long
    ISequence_Length = m_length
End Function

Private Function ISequence_Count(ByVal value As Variant) As Long
    ' Each value appears at most once in a range (values are unique)
    If IContains_Contains(value) Then
        ISequence_Count = 1
    Else
        ISequence_Count = 0
    End If
End Function

Private Function ISequence_Index(ByVal value As Variant, _
                                 Optional ByVal start As Variant, _
                                 Optional ByVal stopAt As Variant) As Long
    ' Normalise start
    Dim startIdx As Long
    If IsMissing(start) Then
        startIdx = 0
    Else
        startIdx = CLng(start)
        If startIdx < 0 Then startIdx = m_length + startIdx
        If startIdx < 0 Then startIdx = 0
    End If

    ' Normalise stop
    Dim stopIdx As Long
    If IsMissing(stopAt) Then
        stopIdx = m_length
    Else
        stopIdx = CLng(stopAt)
        If stopIdx < 0 Then stopIdx = m_length + stopIdx
        If stopIdx > m_length Then stopIdx = m_length
    End If

    ' Range can compute index mathematically (O(1) when unbounded)
    If Not IContains_Contains(value) Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "RangeIterator.Index", "value not in sequence"
    End If

    Dim rawIndex As Long
    rawIndex = (CLng(value) - m_start) \ m_step

    If rawIndex < startIdx Or rawIndex >= stopIdx Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "RangeIterator.Index", "value not in sequence"
    End If

    ISequence_Index = rawIndex
End Function

' ─── ILength — delegates to precomputed length ───

Private Function ILength_Length() As Long
    ILength_Length = m_length
End Function

' ─── IIterable — fast-path iteration via SequenceIterator ───

Private Function IIterable_Iter() As IIterator
    Dim iter As SequenceIterator
    Set iter = New SequenceIterator
    iter.Init Me
    Set IIterable_Iter = iter
End Function
