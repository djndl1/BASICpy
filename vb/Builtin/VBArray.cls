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

Private m_data As Variant   ' backing array (may have spare capacity)
Private m_lb As Long         ' cached LBound(m_data) — stable after Init
Private m_len As Long        ' logical length (elements in use)

Private Sub Class_Initialize()
    m_data = Array()         ' empty array: LBound=0, UBound=-1
    m_lb = 0
    m_len = 0
End Sub

' Initialises from an existing array (copies it).
Public Sub Init(ByRef arr As Variant)
    Ensure.IsTrue IsArray(arr), CommonHResults.TypeMismatch, _
        "VBArray.Init", "arr must be an array"

    ' Guard: undimensioned array (Dim arr() without ReDim) → already empty
    On Error Resume Next
    m_lb = LBound(arr)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    m_len = UBound(arr) - m_lb + 1
    m_data = arr
End Sub

' Returns the number of elements.
Public Function length() As Long
    length = m_len
End Function

' Returns the physical capacity of the backing array.
Private Function Capacity() As Long
    Capacity = UBound(m_data) - m_lb + 1
End Function

' Ensures the backing array can hold at least needed elements.
Private Sub EnsureCapacity(ByVal needed As Long)
    If needed <= Capacity() Then Exit Sub

    Dim cap As Long: cap = Capacity()
    If cap = 0 Then cap = 4
    Do While cap < needed
        cap = cap * 2
    Loop

    Dim tmp() As Variant: tmp = m_data
    ReDim Preserve tmp(m_lb To m_lb + cap - 1)
    m_data = tmp
End Sub


' Returns the element at index (0-based, supports negative indexing).
Public Property Get item(ByVal Index As Long) As Variant
Attribute item.VB_UserMemId = 0
    Builtin.LetSet item, m_data(ArrayUtility.ArrayResolveIndex(Index, m_len, m_lb))
End Property

' Sets the element at index (value types).
Public Property Let item(ByVal Index As Long, ByRef value As Variant)
    m_data(ArrayResolveIndex(Index, m_len, m_lb)) = value
End Property

' Sets the element at index (object references).
Public Property Set item(ByVal Index As Long, ByRef rhs As Variant)
    Set m_data(ArrayResolveIndex(Index, m_len, m_lb)) = rhs
End Property

' ─── IMutableSequence public methods ───

Public Sub Append(ByRef value As Variant)
    EnsureCapacity m_len + 1
    m_data(m_lb + m_len) = value
    m_len = m_len + 1
End Sub

Public Sub Clear()
    m_data = Array()
    m_lb = 0
    m_len = 0
End Sub

' Trims spare capacity and sets logical length (public so Copy can use it).
Public Sub TrimToLength(ByVal targetLen As Long)
    If Capacity() > targetLen * 2 Then
        ReDim Preserve m_data(m_lb To m_lb + targetLen - 1)
    End If
    m_len = targetLen
End Sub

' Returns an independent copy preserving logical length.
Public Function Copy() As Variant
    Dim newArr As VBArray: Set newArr = New VBArray
    If m_len > 0 Then
        newArr.Init m_data
        newArr.TrimToLength m_len
    End If
    Set Copy = newArr
End Function

Public Sub Extend(ByRef iterable As Variant)
    Dim iter As IIterator: Set iter = Builtin.PyIter(iterable)
    Dim item As Variant: Do While Builtin.PyTryNext(iter, item)
        Append item
    Loop
End Sub

Public Sub Insert(ByVal Index As Long, ByRef value As Variant)
    Dim actual As Long: actual = ArrayResolveIndex(Index, m_len, m_lb, True)
    EnsureCapacity m_len + 1

    ' Shift elements right
    Dim i As Long
    For i = m_lb + m_len - 1 To actual Step -1
        m_data(i + 1) = m_data(i)
    Next i

    m_data(actual) = value
    m_len = m_len + 1
End Sub

Public Function Pop(Optional ByVal Index As Long = -1) As Variant
    If m_len = 0 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "VBArray.Pop", "cannot pop from empty array"
    End If

    Dim actual As Long: actual = ArrayResolveIndex(Index, m_len, m_lb)
    Builtin.LetSet Pop, m_data(actual)

    ' Shift elements left (spare capacity provides room)
    Dim i As Long
    For i = actual To m_lb + m_len - 2
        m_data(i) = m_data(i + 1)
    Next i

    m_len = m_len - 1
End Function

Public Sub Remove(ByRef value As Variant)
    Dim writeIdx As Long: writeIdx = m_lb
    Dim readIdx As Long
    Dim found As Boolean

    For readIdx = m_lb To m_lb + m_len - 1
        If m_data(readIdx) = value Then
            found = True
        Else
            If writeIdx <> readIdx Then
                m_data(writeIdx) = m_data(readIdx)
            End If
            writeIdx = writeIdx + 1
        End If
    Next readIdx

    If Not found Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "VBArray.Remove", "value not found"
    End If

    m_len = writeIdx - m_lb
End Sub

Public Sub Reverse()
    ReverseInPlace m_data, m_lb, m_len
End Sub

' ─── IContains — O(n) scan ───

Private Function IContains_Contains(ByVal value As Variant) As Boolean
    IContains_Contains = (ArrayCountOf(m_data, m_lb, m_len, value) > 0)
End Function

' ─── ISequence ───

Private Function ISequence_GetItem(ByVal Index As Long) As Variant
    ISequence_GetItem = item(Index)
End Function

Private Function ISequence_Length() As Long
    ISequence_Length = length()
End Function

Private Function ISequence_Count(ByVal value As Variant) As Long
    ISequence_Count = ArrayCountOf(m_data, m_lb, m_len, value)
End Function

Private Function ISequence_Index(ByVal value As Variant, _
                                 Optional ByVal start As Variant, _
                                 Optional ByVal stopAt As Variant) As Long
    Dim startIdx As Long
    If IsMissing(start) Then startIdx = 0 Else startIdx = CLng(start)
    If startIdx < 0 Then startIdx = m_len + startIdx
    If startIdx < 0 Then startIdx = 0
    
    Dim stopIdx As Long
    If IsMissing(stopAt) Then stopIdx = m_len Else stopIdx = CLng(stopAt)
    If stopIdx < 0 Then stopIdx = m_len + stopIdx
    If stopIdx > m_len Then stopIdx = m_len
    
    ISequence_Index = ArrayIndexOf(m_data, m_lb, m_len, value, startIdx, stopIdx)
    If ISequence_Index = -1 Then
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "VBArray.Index", "value not in sequence"
    End If
End Function

' ─── ILength ───

Private Function ILength_Length() As Long
    ILength_Length = length()
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
