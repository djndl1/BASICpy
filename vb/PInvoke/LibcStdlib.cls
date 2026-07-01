VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcStdlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private m_fmode As UnownedMemoryObject

Private Function EnsureFmode() As UnownedMemoryObject
    If m_fmode Is Nothing Then _
        Set m_fmode = LibcLibrary.ResolveVariable("_fmode")
    Set EnsureFmode = m_fmode
End Function

' extern int _fmode
' Default file translation mode. Values (from LibcTypes.FcntlFlags):
'   O_TEXT, O_BINARY, O_WTEXT
Public Property Get fmode() As Long
    fmode = EnsureFmode.GetAsLong
End Property

Public Property Let fmode(ByVal value As Long)
    EnsureFmode.SetAsLong value
End Property

' ─── Memory allocation ───

' void free(void* memblock)
' Frees a block of memory allocated by malloc, calloc, or realloc.
Public Sub free(ByVal memblock As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("free")
    Dim args(0 To 0) As Variant
    args(0) = memblock
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub

' void* malloc(size_t size)
' Allocates a block of memory. Returns pointer or 0 on failure.
Public Function malloc(ByVal size As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("malloc")
    Dim args(0 To 0) As Variant
    args(0) = size
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    malloc = result
End Function

' void* realloc(void* memblock, size_t size)
' Resizes a memory block. Returns pointer or 0 on failure.
Public Function realloc(ByVal memblock As Long, ByVal size As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("realloc")
    Dim args(0 To 1) As Variant
    args(0) = memblock
    args(1) = size
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    realloc = result
End Function

' ─── Memory copy / move / set ───

' void* memcpy(void* dest, const void* src, size_t count)
' Copies count bytes from src to dest. Does NOT handle overlapping buffers.
Public Sub memcpy(ByVal dest As Long, ByVal src As Long, ByVal count As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("memcpy")
    Dim args(0 To 2) As Variant
    args(0) = dest
    args(1) = src
    args(2) = count
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub

' wchar_t* wmemcpy(wchar_t* dest, const wchar_t* src, size_t count)
' Copies count wchar_t values from src to dest (count is in wchar_t units).
Public Sub wmemcpy(ByVal dest As Long, ByVal src As Long, ByVal count As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wmemcpy", False)
    Dim args(0 To 2) As Variant
    args(0) = dest
    args(1) = src
    args(2) = count
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub

' void* memmove(void* dest, const void* src, size_t count)
' Copies count bytes from src to dest. Handles overlapping buffers.
Public Sub memmove(ByVal dest As Long, ByVal src As Long, ByVal count As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("memmove")
    Dim args(0 To 2) As Variant
    args(0) = dest
    args(1) = src
    args(2) = count
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub

' wchar_t* wmemmove(wchar_t* dest, const wchar_t* src, size_t count)
' Copies count wchar_t values. Handles overlapping buffers.
Public Sub wmemmove(ByVal dest As Long, ByVal src As Long, ByVal count As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wmemmove", False)
    Dim args(0 To 2) As Variant
    args(0) = dest
    args(1) = src
    args(2) = count
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub

' void* memset(void* dest, int c, size_t count)
' Sets count bytes at dest to the value c (truncated to byte).
Public Sub memset(ByVal dest As Long, ByVal c As Long, ByVal count As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("memset")
    Dim args(0 To 2) As Variant
    args(0) = dest
    args(1) = c
    args(2) = count
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub

' wchar_t* wmemset(wchar_t* dest, wchar_t c, size_t count)
' Sets count wchar_t values at dest to c.
Public Sub wmemset(ByVal dest As Long, ByVal c As Long, ByVal count As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wmemset", False)
    Dim args(0 To 2) As Variant
    args(0) = dest
    args(1) = c
    args(2) = count
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub
