VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcIo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' int _chsize(int fd, long size)
' Changes the size of a file. Returns 0 on success, -1 on failure.
Public Function chsize(ByVal fd As Long, ByVal size As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("chsize")
    Dim args(0 To 1) As Variant
    args(0) = fd
    args(1) = size
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    chsize = result
End Function

' int _close(int fd)
' Closes a file descriptor. Returns 0 on success, -1 on failure.
Public Function posix_close(ByVal fd As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("close")
    Dim args(0 To 0) As Variant
    args(0) = fd
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    posix_close = result
End Function

' int _commit(int fd)
' Flushes a file directly to disk. Returns 0 on success, -1 on failure.
Public Function commit(ByVal fd As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("commit")
    Dim args(0 To 0) As Variant
    args(0) = fd
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    commit = result
End Function

' int _dup(int fd)
' Duplicates a file descriptor. Returns new fd or -1 on error.
Public Function dup(ByVal fd As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("dup")
    Dim args(0 To 0) As Variant
    args(0) = fd
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    dup = result
End Function

' int _dup2(int fd1, int fd2)
' Forces fd1 to refer to the same file as fd2. Returns 0 on success, -1 on error.
Public Function dup2(ByVal fd1 As Long, ByVal fd2 As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("dup2")
    Dim args(0 To 1) As Variant
    args(0) = fd1
    args(1) = fd2
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    dup2 = result
End Function

' int _eof(int fd)
' Tests for end of file. Returns 1 if EOF, 0 if not, -1 on error.
Public Function eof(ByVal fd As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("eof")
    Dim args(0 To 0) As Variant
    args(0) = fd
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    eof = result
End Function

' int _wcreat(const wchar_t* filename, int pmode)
' Creates/truncates a file. Returns fd or -1 on error.
' pmode: _S_IWRITE (&H80), _S_IREAD (&H100), or bitwise OR.
Public Function wcreat(ByVal filename As String, ByVal pmode As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("wcreat")
    Dim args(0 To 1) As Variant
    args(0) = StrPtr(filename)
    args(1) = pmode
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wcreat = result
End Function

' ─── File size / find ───

' long _filelength(int fd)
' Gets the length of a file in bytes. Returns file length or -1L on error.
' Works for files up to 2 GB. For larger files, use filelengthi64.
Public Function filelength(ByVal fd As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_filelength", False)
    Dim args(0 To 0) As Variant
    args(0) = fd
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    filelength = result
End Function

' __int64 _filelengthi64(int fd)
' Gets the length of a file in bytes (64-bit). Returns the value as Currency
' (VB6's only 8-byte numeric type — the raw __int64 bit pattern is preserved).
' To display the actual integer, multiply by 10000:
'   Dim len64 As Currency: len64 = LibcIo.filelengthi64(fd)
'   Debug.Print CCur(len64 * 10000)
Public Function filelengthi64(ByVal fd As Long) As Currency
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_filelengthi64", False)
    Dim args(0 To 0) As Variant
    args(0) = fd
    Dim result As Variant
    fn.CallFunc vbCurrency, result, args
    filelengthi64 = result
End Function

' int _findclose(intptr_t handle)
' Closes a file search handle returned by _findfirst / _findnext.
' Returns 0 on success, -1 on error.
Public Function findclose(ByVal handle As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_findclose", False)
    Dim args(0 To 0) As Variant
    args(0) = handle
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    findclose = result
End Function

' int _wopen(const wchar_t* filename, int oflag [, int pmode])
' Opens a file using a Unicode filename. Returns fd or -1 on error.
' oflag: combo of FcntlFlags (O_RDONLY, O_WRONLY, O_RDWR, O_CREAT, O_BINARY, ...)
' pmode (optional): _S_IWRITE, _S_IREAD from io.h (only used when O_CREAT is set)
Public Function wopen(ByVal filename As String, ByVal oflag As Long, _
                      Optional ByVal pmode As Long = 0) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wopen", False)
    Dim args(0 To 2) As Variant
    args(0) = StrPtr(filename)
    args(1) = oflag
    args(2) = pmode
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wopen = result
End Function
