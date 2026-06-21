VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IContains
Implements ISequence
Implements ILength
Implements IIterable
Implements IMutableSequence

Private m_data As Variant   ' holds a 0-based array of Variants

Private Sub Class_Initialize()
    m_data = Array()         ' empty array: LBound=0, UBound=-1
End Sub

' Initialises from an existing array (copies it).
Public Sub Init(ByRef arr As Variant)
    Ensure.IsTrue IsArray(arr), CommonHResults.TypeMismatch, _
        "VBArray.Init", "arr must be an array"
    m_data = arr
End Sub

' Returns the number of elements.
Public Function Length() As Long
    Length = UBound(m_data) + 1
End Function

' Returns the element at index (0-based, supports negative indexing).
Public Property Get item(ByVal Index As Long) As Variant
Attribute item.VB_UserMemId = 0
    If Index < 0 Then Index = Length() + Index
    If Index < 0 Or Index > UBound(m_data) Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "VBArray.Item", "index out of range"
    End If
    Builtin.LetSet item, m_data(Index)
End Property

' Sets the element at index (value types).
Public Property Let item(ByVal Index As Long, ByRef value As Variant)
    If Index < 0 Then Index = Length() + Index
    If Index < 0 Or Index > UBound(m_data) Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "VBArray.Item", "index out of range"
    End If
    m_data(Index) = value
End Property

' Sets the element at index (object references).
Public Property Set item(ByVal Index As Long, ByRef rhs As Variant)
    If Index < 0 Then Index = Length() + Index
    If Index < 0 Or Index > UBound(m_data) Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "VBArray.Item", "index out of range"
    End If
    Set m_data(Index) = rhs
End Property

' ─── IMutableSequence public methods ───

Public Sub Append(ByRef value As Variant)
    Dim sz As Long: sz = UBound(m_data)
    Dim tmp() As Variant: tmp = m_data
    ReDim Preserve tmp(0 To sz + 1)
    m_data = tmp
    m_data(sz + 1) = value
End Sub

Public Sub Clear()
    m_data = Array()
End Sub

Public Function Copy() As Variant
    Dim newArr As VBArray
    Set newArr = New VBArray
    newArr.Init m_data
    Set Copy = newArr
End Function

Public Sub Extend(ByRef iterable As Variant)
    Dim iter As IIterator
    Set iter = Builtin.PyIter(iterable)
    Dim item As Variant
    Do While Builtin.PyTryNext(iter, item)
        Append item
    Loop
End Sub

Public Sub Insert(ByVal Index As Long, ByRef value As Variant)
    Dim sz As Long: sz = UBound(m_data)
    If Index < 0 Then Index = sz + 1 + Index
    If Index < 0 Or Index > sz + 1 Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "VBArray.Insert", "index out of range"
    End If
    
    ' Grow by one
    Dim tmp() As Variant: tmp = m_data
    ReDim Preserve tmp(0 To sz + 1)
    
    ' Shift elements right from index to end
    Dim i As Long
    For i = sz To Index Step -1
        tmp(i + 1) = tmp(i)
    Next i
    
    tmp(Index) = value
    m_data = tmp
End Sub

Public Function Pop(Optional ByVal Index As Long = -1) As Variant
    Dim sz As Long: sz = UBound(m_data)
    If sz < 0 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "VBArray.Pop", "cannot pop from empty array"
    End If
    
    If Index < 0 Then Index = sz + 1 + Index
    If Index < 0 Or Index > sz Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "VBArray.Pop", "index out of range"
    End If
    
    Builtin.LetSet Pop, m_data(Index)
    
    ' Shift elements left
    Dim tmp() As Variant: tmp = m_data
    Dim i As Long
    For i = Index To sz - 1
        tmp(i) = tmp(i + 1)
    Next i
    
    If sz > 0 Then
        ReDim Preserve tmp(0 To sz - 1)
    Else
        tmp = Array()
    End If
    m_data = tmp
End Function

Public Sub Remove(ByRef value As Variant)
    Dim sz As Long: sz = UBound(m_data)
    Dim i As Long
    For i = 0 To sz
        If m_data(i) = value Then
            Pop i
            Exit Sub
        End If
    Next i
    
    Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
        "VBArray.Remove", "value not found"
End Sub

Public Sub Reverse()
    Dim tmp() As Variant: tmp = m_data
    Dim i As Long, j As Long
    i = 0: j = UBound(tmp)
    Dim t As Variant
    While i < j
        t = tmp(i)
        tmp(i) = tmp(j)
        tmp(j) = t
        i = i + 1
        j = j - 1
    Wend
    m_data = tmp
End Sub

' ─── IContains — O(n) scan ───

Private Function IContains_Contains(ByVal value As Variant) As Boolean
    Dim i As Long
    For i = 0 To UBound(m_data)
        If m_data(i) = value Then
            IContains_Contains = True
            Exit Function
        End If
    Next i
    IContains_Contains = False
End Function

' ─── ISequence ───

Private Function ISequence_GetItem(ByVal Index As Long) As Variant
    ISequence_GetItem = item(Index)
End Function

Private Function ISequence_Length() As Long
    ISequence_Length = Length()
End Function

Private Function ISequence_Count(ByVal value As Variant) As Long
    Dim i As Long
    For i = 0 To UBound(m_data)
        If m_data(i) = value Then ISequence_Count = ISequence_Count + 1
    Next i
End Function

Private Function ISequence_Index(ByVal value As Variant, _
                                 Optional ByVal start As Variant, _
                                 Optional ByVal stopAt As Variant) As Long
    Dim sz As Long: sz = UBound(m_data)
    
    Dim startIdx As Long
    If IsMissing(start) Then startIdx = 0 Else startIdx = CLng(start)
    If startIdx < 0 Then startIdx = sz + 1 + startIdx
    If startIdx < 0 Then startIdx = 0
    
    Dim stopIdx As Long
    If IsMissing(stopAt) Then stopIdx = sz + 1 Else stopIdx = CLng(stopAt)
    If stopIdx < 0 Then stopIdx = sz + 1 + stopIdx
    If stopIdx > sz + 1 Then stopIdx = sz + 1
    
    Dim i As Long
    For i = startIdx To stopIdx - 1
        If m_data(i) = value Then
            ISequence_Index = i
            Exit Function
        End If
    Next i
    
    Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
        "VBArray.Index", "value not in sequence"
End Function

' ─── ILength ───

Private Function ILength_Length() As Long
    ILength_Length = Length()
End Function

' ─── IIterable — fast-path iteration via SequenceIterator ───

Private Function IIterable_Iter() As IIterator
    Dim iter As SequenceIterator
    Set iter = New SequenceIterator
    iter.Init Me
    Set IIterable_Iter = iter
End Function

' ─── IMutableSequence interface implementation ───

Private Sub IMutableSequence_Append(ByRef value As Variant)
    Append value
End Sub

Private Sub IMutableSequence_Clear()
    Clear
End Sub

Private Function IMutableSequence_Copy() As Variant
    Set IMutableSequence_Copy = Copy()
End Function

Private Sub IMutableSequence_Extend(ByRef iterable As Variant)
    Extend iterable
End Sub

Private Sub IMutableSequence_Insert(ByVal Index As Long, ByRef value As Variant)
    Insert Index, value
End Sub

Private Function IMutableSequence_Pop(Optional ByVal Index As Long = -1) As Variant
    IMutableSequence_Pop = Pop(Index)
End Function

Private Sub IMutableSequence_Remove(ByRef value As Variant)
    Remove value
End Sub

Private Sub IMutableSequence_Reverse()
    Reverse
End Sub
