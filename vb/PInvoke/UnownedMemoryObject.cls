VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnownedMemoryObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_ptr As Long

' Initializes this wrapper with a pointer to the DLL variable.
Public Sub Init(ByVal ptr As Long)
    m_ptr = ptr
End Sub

' Returns the raw pointer to the DLL variable.
Public Property Get ptr() As Long
    ptr = m_ptr
End Property

' Whether this wrapper holds a valid (non-null) pointer.
Public Property Get IsInitialized() As Boolean
    IsInitialized = (m_ptr <> 0)
End Property

' ─── Read ───

Public Function GetAsByte() As Byte
    If m_ptr = 0 Then Exit Function
    Dim val As Byte
    CopyMemory val, ByVal m_ptr, 1
    GetAsByte = val
End Function

Public Function GetAsInteger() As Integer
    If m_ptr = 0 Then Exit Function
    Dim val As Integer
    CopyMemory val, ByVal m_ptr, 2
    GetAsInteger = val
End Function

Public Function GetAsLong() As Long
    If m_ptr = 0 Then Exit Function
    Dim val As Long
    CopyMemory val, ByVal m_ptr, 4
    GetAsLong = val
End Function

Public Function GetAsSingle() As Single
    If m_ptr = 0 Then Exit Function
    Dim val As Single
    CopyMemory val, ByVal m_ptr, 4
    GetAsSingle = val
End Function

Public Function GetAsDouble() As Double
    If m_ptr = 0 Then Exit Function
    Dim val As Double
    CopyMemory val, ByVal m_ptr, 8
    GetAsDouble = val
End Function

' Reads bytes from the variable into a caller-provided Byte array.
' The transfer length is determined by UBound(data) - LBound(data) + 1.
Public Sub GetAsByteArray(ByRef data() As Byte)
    If m_ptr = 0 Then Exit Sub
    Dim lenBytes As Long
    On Error Resume Next
    lenBytes = UBound(data) - LBound(data) + 1
    If Err.Number <> 0 Or lenBytes < 1 Then Exit Sub
    Err.Clear
    On Error GoTo 0
    CopyMemory data(LBound(data)), ByVal m_ptr, lenBytes
End Sub

Public Function GetAsString() As String
    GetAsString = NativeUtility.NullTerminatedWStringToBStr(m_ptr)
End Function

' ─── Write ───

Public Sub SetAsByte(ByVal value As Byte)
    If m_ptr = 0 Then Exit Sub
    CopyMemory ByVal m_ptr, value, 1
End Sub

Public Sub SetAsInteger(ByVal value As Integer)
    If m_ptr = 0 Then Exit Sub
    CopyMemory ByVal m_ptr, value, 2
End Sub

Public Sub SetAsLong(ByVal value As Long)
    If m_ptr = 0 Then Exit Sub
    CopyMemory ByVal m_ptr, value, 4
End Sub

Public Sub SetAsSingle(ByVal value As Single)
    If m_ptr = 0 Then Exit Sub
    CopyMemory ByVal m_ptr, value, 4
End Sub

Public Sub SetAsDouble(ByVal value As Double)
    If m_ptr = 0 Then Exit Sub
    CopyMemory ByVal m_ptr, value, 8
End Sub

' Writes data from a Byte array into the variable.
Public Sub SetAsByteArray(ByRef data() As Byte)
    If m_ptr = 0 Then Exit Sub
    Dim lenBytes As Long
    On Error Resume Next
    lenBytes = UBound(data) - LBound(data) + 1
    If Err.Number <> 0 Or lenBytes < 1 Then Exit Sub
    Err.Clear
    On Error GoTo 0
    CopyMemory ByVal m_ptr, data(LBound(data)), lenBytes
End Sub

' ─── String operations ───
Public Sub SetAsString(ByRef value As String)
    If m_ptr = 0 Then Exit Sub
    NativeUtility.CopyStringToNullTerminatedWString m_ptr, value
End Sub
