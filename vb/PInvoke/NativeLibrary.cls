VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NativeLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Wraps a native DLL handle with automatic cleanup via Class_Terminate.
' Usage:
'   Dim lib As New NativeLibrary
'   lib.Load "user32.dll"
'   Dim fn As CDeclLibraryFunction: Set fn = lib.GetProcAddress("MessageBoxW")
'   ' ... use fn.CallFunc or fn.Ptr with CDeclCall ...
'   ' lib is automatically freed when it goes out of scope

Private m_hModule As Long

' Loads the specified DLL into the process address space.
' Raises if the DLL cannot be found or loaded.
Public Sub Load(ByVal dllName As String)
    ' Unload previous library if one is already loaded
    If m_hModule <> 0 Then
        DispCall.FreeLibrary m_hModule
        m_hModule = 0
    End If
    
    m_hModule = DispCall.LoadLibraryW(StrPtr(dllName))
    If m_hModule = 0 Then
        Err.Raise vbObjectError + 1024, "NativeLibrary.Load", _
            "Failed to load library: " & dllName
    End If
End Sub

' Returns a CDeclLibraryFunction wrapping an exported function in the loaded DLL.
' Raises if the function is not found.
Public Function GetProcAddress(ByVal procName As String) As CDeclLibraryFunction
    If m_hModule = 0 Then
        Err.Raise vbObjectError + 1025, "NativeLibrary.GetProcAddress", _
            "No library is currently loaded"
    End If
    
    Dim fnPtr As Long
    fnPtr = DispCall.GetProcAddress(m_hModule, procName)
    If fnPtr = 0 Then
        Err.Raise vbObjectError + 1026, "NativeLibrary.GetProcAddress", _
            "Function not found: " & procName
    End If
    
    Dim fn As New CDeclLibraryFunction
    fn.Init fnPtr
    Set GetProcAddress = fn
End Function

' Returns the raw HMODULE handle for direct use with other API calls.
Public Property Get Handle() As Long
    Handle = m_hModule
End Property

' Automatically unloads the library when the object is destroyed.
Private Sub Class_Terminate()
    If m_hModule <> 0 Then
        DispCall.FreeLibrary m_hModule
        m_hModule = 0
    End If
End Sub
