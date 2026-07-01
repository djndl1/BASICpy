VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OleMemoryObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements UnownedMemoryObject

' OleMemoryObject wraps an UnownedMemoryObject
' It delegates all operations to the wrapped UnownedMemoryObject
' Only Class_Terminate is its own implementation - it frees the memory
Private m_wrapped As UnownedMemoryObject

' Constructor: wraps an existing UnownedMemoryObject
Public Sub Class_Initialize()
    Set m_wrapped = New UnownedMemoryObject
End Sub

' Destructor: frees the memory using CoTaskMemFree
Private Sub Class_Terminate()
    If Not (m_wrapped Is Nothing) Then
        ' Free the memory using CoTaskMemFree
        RtlWinApi.CoTaskMemFree m_wrapped.ptr
        Set m_wrapped = Nothing
    End If
End Sub

' ─── UnownedMemoryObject interface implementation ───
Private Sub UnownedMemoryObject_Init(ByVal ptr As Long)
    m_wrapped.Init ptr
End Sub

Private Property Get UnownedMemoryObject_ptr() As Long
    UnownedMemoryObject_ptr = m_wrapped.ptr
End Property

Private Property Get UnownedMemoryObject_IsInitialized() As Boolean
    UnownedMemoryObject_IsInitialized = m_wrapped.IsInitialized
End Property

Private Function UnownedMemoryObject_GetAsByte() As Byte
    UnownedMemoryObject_GetAsByte = m_wrapped.GetAsByte
End Function

Private Function UnownedMemoryObject_GetAsInteger() As Integer
    UnownedMemoryObject_GetAsInteger = m_wrapped.GetAsInteger
End Function

Private Function UnownedMemoryObject_GetAsLong() As Long
    UnownedMemoryObject_GetAsLong = m_wrapped.GetAsLong
End Function

Private Function UnownedMemoryObject_GetAsSingle() As Single
    UnownedMemoryObject_GetAsSingle = m_wrapped.GetAsSingle
End Function

Private Function UnownedMemoryObject_GetAsDouble() As Double
    UnownedMemoryObject_GetAsDouble = m_wrapped.GetAsDouble
End Function

Private Sub UnownedMemoryObject_GetAsByteArray(ByRef data() As Byte)
    m_wrapped.GetAsByteArray data
End Sub

Private Function UnownedMemoryObject_GetAsString() As String
    UnownedMemoryObject_GetAsString = m_wrapped.GetAsString
End Function

Private Sub UnownedMemoryObject_SetAsByte(ByVal value As Byte)
    m_wrapped.SetAsByte value
End Sub

Private Sub UnownedMemoryObject_SetAsInteger(ByVal value As Integer)
    m_wrapped.SetAsInteger value
End Sub

Private Sub UnownedMemoryObject_SetAsLong(ByVal value As Long)
    m_wrapped.SetAsLong value
End Sub

Private Sub UnownedMemoryObject_SetAsSingle(ByVal value As Single)
    m_wrapped.SetAsSingle value
End Sub

Private Sub UnownedMemoryObject_SetAsDouble(ByVal value As Double)
    m_wrapped.SetAsDouble value
End Sub

Private Sub UnownedMemoryObject_SetAsByteArray(ByRef data() As Byte)
    m_wrapped.SetAsByteArray data
End Sub

Private Sub UnownedMemoryObject_SetAsString(ByRef value As String)
    m_wrapped.SetAsString value
End Sub

' ─── Public interface (direct access) ───
' Initializes this wrapper with a pointer to the DLL variable.
Public Sub Init(ByVal ptr As Long)
    m_wrapped.Init ptr
End Sub

' Returns the raw pointer to the DLL variable.
Public Property Get ptr() As Long
    ptr = m_wrapped.ptr
End Property

' Whether this wrapper holds a valid (non-null) pointer.
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_wrapped.IsInitialized
End Property

' ─── Read ───
Public Function GetAsByte() As Byte
    GetAsByte = m_wrapped.GetAsByte
End Function

Public Function GetAsInteger() As Integer
    GetAsInteger = m_wrapped.GetAsInteger
End Function

Public Function GetAsLong() As Long
    GetAsLong = m_wrapped.GetAsLong
End Function

Public Function GetAsSingle() As Single
    GetAsSingle = m_wrapped.GetAsSingle
End Function

Public Function GetAsDouble() As Double
    GetAsDouble = m_wrapped.GetAsDouble
End Function

' Reads bytes from the variable into a caller-provided Byte array.
Public Sub GetAsByteArray(ByRef data() As Byte)
    m_wrapped.GetAsByteArray data
End Sub

Public Function GetAsString() As String
    GetAsString = m_wrapped.GetAsString
End Function

' ─── Write ───
Public Sub SetAsByte(ByVal value As Byte)
    m_wrapped.SetAsByte value
End Sub

Public Sub SetAsInteger(ByVal value As Integer)
    m_wrapped.SetAsInteger value
End Sub

Public Sub SetAsLong(ByVal value As Long)
    m_wrapped.SetAsLong value
End Sub

Public Sub SetAsSingle(ByVal value As Single)
    m_wrapped.SetAsSingle value
End Sub

Public Sub SetAsDouble(ByVal value As Double)
    m_wrapped.SetAsDouble value
End Sub

' Writes data from a Byte array into the variable.
Public Sub SetAsByteArray(ByRef data() As Byte)
    m_wrapped.SetAsByteArray data
End Sub

' ─── String operations ───
Public Sub SetAsString(ByRef value As String)
    m_wrapped.SetAsString value
End Sub