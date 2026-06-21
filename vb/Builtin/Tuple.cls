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

' Initialises the tuple from an array. Copies the array to enforce immutability.
' Raises TypeMismatch if value is not an array.
Public Sub Init(ByRef arr As Variant)
    Ensure.IsTrue IsArray(arr), CommonHResults.TypeMismatch, _
        "Tuple.Init", "arr must be an array"
    
    m_data = arr
    m_length = UBound(m_data) - LBound(m_data) + 1
End Sub

' Returns the number of elements.
Public Function Length() As Long
    Length = m_length
End Function

' Returns the element at the given index (0-based).
' Supports negative indexing: -1 is the last element.
Public Function GetItem(ByVal Index As Long) As Variant
    If Index < 0 Then
        Index = m_length + Index
    End If
    
    If Index < 0 Or Index >= m_length Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "Tuple.GetItem", "index out of range"
    End If
    
    Builtin.LetSet GetItem, m_data(LBound(m_data) + Index)
End Function

' Returns the number of occurrences of value.
Public Function count(ByVal value As Variant) As Long
    Dim i As Long
    Dim lb As Long: lb = LBound(m_data)
    Dim ub As Long: ub = UBound(m_data)
    
    For i = lb To ub
        If m_data(i) = value Then count = count + 1
    Next i
End Function

' Returns the index of the first occurrence of value within [start, stop).
' Raises if value is not found.
Public Function Index(ByVal value As Variant, _
                      Optional ByVal start As Variant, _
                      Optional ByVal stopAt As Variant) As Long
    Dim startIdx As Long
    If IsMissing(start) Then
        startIdx = 0
    Else
        startIdx = CLng(start)
        If startIdx < 0 Then startIdx = m_length + startIdx
        If startIdx < 0 Then startIdx = 0
    End If
    
    Dim stopIdx As Long
    If IsMissing(stopAt) Then
        stopIdx = m_length
    Else
        stopIdx = CLng(stopAt)
        If stopIdx < 0 Then stopIdx = m_length + stopIdx
        If stopIdx > m_length Then stopIdx = m_length
    End If
    
    Dim i As Long
    Dim lb As Long: lb = LBound(m_data) + startIdx
    Dim ub As Long: ub = LBound(m_data) + stopIdx - 1
    If ub > UBound(m_data) Then ub = UBound(m_data)
    
    For i = lb To ub
        If m_data(i) = value Then
            Index = i - LBound(m_data)
            Exit Function
        End If
    Next i
    
    Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
        "Tuple.Index", "value not in sequence"
End Function

' Default Item property — enables tupleObj(index) syntax.
Public Property Get item(ByVal Index As Long) As Variant
Attribute item.VB_UserMemId = 0
    Builtin.LetSet item, GetItem(Index)
End Property

' ─── IContains — O(n) scan ───

Private Function IContains_Contains(ByVal value As Variant) As Boolean
    Dim i As Long
    For i = LBound(m_data) To UBound(m_data)
        If m_data(i) = value Then
            IContains_Contains = True
            Exit Function
        End If
    Next i
    IContains_Contains = False
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
