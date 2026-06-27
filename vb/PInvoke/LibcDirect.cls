VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' int _wchdir(const wchar_t* dirname)
' Changes the current working directory. Returns 0 on success, -1 on failure.
Public Function wchdir(ByVal dirname As String) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("wchdir")
    Dim args(0 To 0) As Variant
    args(0) = StrPtr(dirname)
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wchdir = result
End Function

' int _chdrive(int drive)
' Changes the current drive (1=A, 2=B, 3=C, ...). Returns 0 on success, -1 on failure.
Public Function chdrive(ByVal drive As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("chdrive")
    Dim args(0 To 0) As Variant
    args(0) = drive
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    chdrive = result
End Function

' wchar_t* _wgetcwd(wchar_t* buffer, int maxlen)
' Gets the current working directory as a wide string.
' If buffer is 0, allocates with malloc (caller must free).
' Returns pointer to buffer or 0 on failure.
'   Dim buf(0 To 511) As Byte     ' 256 wchar_t
'   Dim ptr As Long: ptr = LibcDirect.wgetcwd(VarPtr(buf(0)), 256)
Public Function wgetcwd(ByVal buffer As Long, ByVal maxlen As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wgetcwd", False)
    Dim args(0 To 1) As Variant
    args(0) = buffer
    args(1) = maxlen
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wgetcwd = result
End Function
