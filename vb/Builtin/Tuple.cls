VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IContains
Implements ISequence
Implements ILength
Implements IIterable

Private m_data As Variant          ' internal array copy (enforces immutability)
Private m_length As Long           ' precomputed length for O(1) access
Private m_lb As Long               ' cached LBound for O(1) indexing

' Initialises the tuple from an array. Copies the array to enforce immutability.
' Raises TypeMismatch if value is not an array.
Public Sub Init(ByRef arr As Variant)
    Ensure.IsTrue IsArray(arr), CommonHResults.TypeMismatch, _
        "Tuple.Init", "arr must be an array"

    m_data = arr

    ' Guard undimensioned arrays: IsArray passes but LBound/UBound crash
    On Error Resume Next
    m_lb = LBound(m_data)
    m_length = UBound(m_data) - m_lb + 1
    If Err.Number <> 0 Then
        m_length = 0
        m_lb = 0
        m_data = Array()
    End If
    On Error GoTo 0
End Sub

' Returns the number of elements.
Public Function Length() As Long
    Length = m_length
End Function

' Returns the element at the given index (0-based).
' Supports negative indexing: -1 is the last element.
Public Function GetItem(ByVal Index As Long) As Variant
    Builtin.LetSet GetItem, m_data(ArrayResolveIndex(Index, m_length, m_lb))
End Function

' Returns the number of occurrences of value.
Public Function count(ByVal value As Variant) As Long
    count = ArrayCountOf(m_data, m_lb, m_length, value)
End Function

' Returns the index of the first occurrence of value within [start, stop).
' Raises if value is not found.
Public Function Index(ByVal value As Variant, _
                      Optional ByVal start As Variant, _
                      Optional ByVal stopAt As Variant) As Long
    ' startIdx defaults to 0 when start is missing
    Dim startIdx As Long
    If Not IsMissing(start) Then startIdx = CLng(start)
    If startIdx < 0 Then startIdx = m_length + startIdx
    If startIdx < 0 Then startIdx = 0

    ' stopIdx defaults to m_length when stopAt is missing
    Dim stopIdx As Long
    stopIdx = m_length
    If Not IsMissing(stopAt) Then stopIdx = CLng(stopAt)
    If stopIdx < 0 Then stopIdx = m_length + stopIdx
    If stopIdx > m_length Then stopIdx = m_length

    Index = ArrayIndexOf(m_data, m_lb, m_length, value, startIdx, stopIdx)
    If Index = -1 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Tuple.Index", "value not in sequence"
    End If
End Function

' Default Item property — enables tupleObj(index) syntax.
Public Property Get Item(ByVal Index As Long) As Variant
Attribute item.VB_UserMemId = 0
    Builtin.LetSet item, GetItem(Index)
End Property

' ─── IContains — O(n) scan ───

Private Function IContains_Contains(ByVal value As Variant) As Boolean
    IContains_Contains = (ArrayCountOf(m_data, m_lb, m_length, value) > 0)
End Function

' ─── ISequence ───

Private Function ISequence_GetItem(ByVal Index As Long) As Variant
    ISequence_GetItem = GetItem(Index)
End Function

Private Function ISequence_Length() As Long
    ISequence_Length = m_length
End Function

Private Function ISequence_Count(ByVal value As Variant) As Long
    ISequence_Count = count(value)
End Function

Private Function ISequence_Index(ByVal value As Variant, _
                                 Optional ByVal start As Variant, _
                                 Optional ByVal stopAt As Variant) As Long
    ISequence_Index = Index(value, start, stopAt)
End Function

' ─── ILength ───

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