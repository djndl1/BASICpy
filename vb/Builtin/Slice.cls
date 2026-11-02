VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Slice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ─── Internal state (Empty = Python's None) ───

Private m_start As Variant
Private m_stop As Variant
Private m_step As Variant

' ─── Constructor ───

' Initialises the Slice with start, stop, step.
' Each may be Empty (Python's None), a numeric value, or an IIndex object.
Public Sub Init(ByRef start As Variant, ByRef stopIndex As Variant, ByRef step As Variant)
    m_start = start
    m_stop = stopIndex
    m_step = step
End Sub

' ─── Read-only properties ───

' Returns the start index (may be Empty for None).
Public Property Get SliceStart() As Variant
    SliceStart = m_start
End Property

' Returns the stop index (may be Empty for None).
Public Property Get SliceStop() As Variant
    SliceStop = m_stop
End Property

' Returns the step value (may be Empty for None).
Public Property Get SliceStep() As Variant
    SliceStep = m_step
End Property

' ─── Helpers ───

' Resolves a single slice index value to a Long.
' Supports IIndex protocol (Python's __index__ equivalent).
' Caller guarantees value is not Empty (None).
Private Function ResolveIndex(ByRef value As Variant) As Long
    If IsObject(value) Then
        If TypeOf value Is IIndex Then
            Dim idx As IIndex: Set idx = value
            ResolveIndex = CLng(idx.Index())
        Else
            Err.Raise ErrorCodes.TypeMismatch, "Slice.ResolveIndex", "slice indices must be integers or IIndex objects"
        End If
    Else
        ResolveIndex = CLng(value)
    End If
End Function

' ─── Public API ───

' Matches CPython slice.indices(length) — _PySlice_GetLongIndices
' Returns a 3-element Variant array: (start, stop, step).
' Clamps out-of-bounds indices per Python semantics.
' Raises if step is zero.
Public Function Indices(ByVal length As Long) As Variant
    Dim result(0 To 2) As Long
    Dim s As Long, st As Long, sp As Long
    Dim stepIsNegative As Boolean
    Dim lower As Long, upper As Long

    ' Resolve step
    If IsEmpty(m_step) Then
        sp = 1
        stepIsNegative = False
    Else
        sp = ResolveIndex(m_step)
        If sp = 0 Then
            Err.Raise ErrorCodes.InvalidProcedureCall, "Slice.Indices", "slice step cannot be zero"
        End If
        stepIsNegative = (sp < 0)
    End If

    ' Compute bounds per CPython _PySlice_GetLongIndices
    If stepIsNegative Then
        lower = -1
        upper = length - 1
    Else
        lower = 0
        upper = length
    End If

    ' Resolve start
    If IsEmpty(m_start) Then
        If stepIsNegative Then s = upper Else s = lower
    Else
        s = ResolveIndex(m_start)
        If s < 0 Then
            s = s + length
            If s < lower Then s = lower
        Else
            If s > upper Then s = upper
        End If
    End If

    ' Resolve stop
    If IsEmpty(m_stop) Then
        If stepIsNegative Then st = lower Else st = upper
    Else
        st = ResolveIndex(m_stop)
        If st < 0 Then
            st = st + length
            If st < lower Then st = lower
        Else
            If st > upper Then st = upper
        End If
    End If

    result(0) = s
    result(1) = st
    result(2) = sp
    Indices = result
End Function
