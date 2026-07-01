VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Internal singleton that wraps msvcrt.dll (Microsoft C Runtime) for CDECL
' function resolution. All public Libc* classes (LibcMath, LibcErrno, etc.)
' resolve their function pointers through this class, sharing a single
' NativeLibrary instance and a cached function-pointer Collection.

Private m_lib As NativeLibrary
Private m_cache As Collection

Private Sub Class_Initialize()
    Set m_lib = New NativeLibrary
    m_lib.Load "msvcrt.dll"
    Set m_cache = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_cache = Nothing
    Set m_lib = Nothing
End Sub

' Shared resolver: cache lookup → GetProcAddress → wraps result.
' asVariable=False returns CDeclLibraryFunction (for ResolveFunction).
' asVariable=True  returns UnownedMemoryObject  (for ResolveVariable).
Private Function ResolveRaw(ByVal name As String, _
                            Optional ByVal mangle As Boolean = True, _
                            Optional ByVal asVariable As Boolean = False) As Object
    On Error Resume Next
    Set ResolveRaw = m_cache(name)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    On Error GoTo 0
    
    Dim lookupName As String
    If mangle Then lookupName = DispCall.CDeclMangle(name) Else lookupName = name
    
    Dim raw As Object
    Set raw = m_lib.GetProcAddress(lookupName)
    If raw Is Nothing Then Exit Function
    
    If asVariable Then
        Dim var As UnownedMemoryObject
        Set var = New UnownedMemoryObject
        var.Init raw.Ptr
        m_cache.Add var, name
        Set ResolveRaw = var
    Else
        m_cache.Add raw, name
        Set ResolveRaw = raw
    End If
End Function

' Resolves a C runtime function by name.
' Set mangle=False for names where the underscore is part of the export
' (e.g., "_wfopen", "_cgets", "_aligned_free").
Public Function ResolveFunction(ByVal name As String, _
                                Optional ByVal mangle As Boolean = True) As CDeclLibraryFunction
    Set ResolveFunction = ResolveRaw(name, mangle, False)
End Function

' Resolves a C runtime data export (variable, not a function) by name.
' Caches the UnownedMemoryObject directly.
Public Function ResolveVariable(ByVal name As String, _
                                Optional ByVal mangle As Boolean = True) As UnownedMemoryObject
    Set ResolveVariable = ResolveRaw(name, mangle, True)
End Function

' Returns the underlying NativeLibrary for advanced or uncached use.
Public Property Get NativeLib() As NativeLibrary
    Set NativeLib = m_lib
End Property