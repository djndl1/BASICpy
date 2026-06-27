VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcStdio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' FILE* _wfopen(const wchar_t* filename, const wchar_t* mode)
' Opens a file with a Unicode filename and mode string.
' Returns a FILE* handle (Long), or 0 on failure.
' Note: export name is "_wfopen" (underscore is part of the name, not CDECL decoration),
'       so mangle is set to False.
Public Function wfopen(ByVal filename As String, ByVal mode As String) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wfopen", False)
    Dim args(0 To 1) As Variant
    args(0) = StrPtr(filename)   ' LPCWSTR
    args(1) = StrPtr(mode)       ' LPCWSTR
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wfopen = result
End Function

' int fclose(FILE* stream)
' Closes a file stream. Returns 0 on success, EOF on failure.
Public Function fclose(ByVal stream As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("fclose")
    Dim args(0 To 0) As Variant
    args(0) = stream
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fclose = result
End Function

' int _fcloseall(void)
' Closes all open file streams (except stdin/stdout/stderr). Returns total
' streams closed, or EOF on error.
Public Function fcloseall() As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_fcloseall", False)
    Dim args(0 To 0) As Variant
    args(0) = 0&  ' dummy — _fcloseall() takes no args
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fcloseall = result
End Function

' FILE* _wfdopen(int fd, const wchar_t* mode)
' Associates a file descriptor with a wide-character file stream.
' Returns a FILE* handle (Long), or 0 on failure.
Public Function wfdopen(ByVal fd As Long, ByVal mode As String) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wfdopen", False)
    Dim args(0 To 1) As Variant
    args(0) = fd
    args(1) = StrPtr(mode)
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wfdopen = result
End Function

' int _fileno(FILE* stream)
' Gets the file descriptor associated with a stream. Returns fd or -1 on error.
Public Function fileno(ByVal stream As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_fileno", False)
    Dim args(0 To 0) As Variant
    args(0) = stream
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fileno = result
End Function

' int _fgetchar(void)
' Reads the next character from stdin (ANSI). Returns the character as an int,
' or EOF on error/end-of-file.
Public Function fgetchar() As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_fgetchar", False)
    Dim args(0 To 0) As Variant
    args(0) = 0&  ' dummy — _fgetchar() takes no args
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fgetchar = result
End Function

' wint_t _fgetwchar(void)
' Reads the next wide character from stdin. Returns the character as a Long,
' or WEOF on error/end-of-file.
Public Function fgetwchar() As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_fgetwchar", False)
    Dim args(0 To 0) As Variant
    args(0) = 0&  ' dummy — _fgetwchar() takes no args
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fgetwchar = result
End Function

' int _flushall(void)
' Flushes all streams (writes all buffers to files). Returns number of
' streams flushed.
Public Function flushall() As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_flushall", False)
    Dim args(0 To 0) As Variant
    args(0) = 0&  ' dummy — _flushall() takes no args
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    flushall = result
End Function

' int _fputchar(int c)
' Writes a character to stdout (ANSI). Returns the character written, or EOF on error.
Public Function fputchar(ByVal c As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_fputchar", False)
    Dim args(0 To 0) As Variant
    args(0) = c
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fputchar = result
End Function

' wint_t _fputwchar(wchar_t c)
' Writes a wide character to stdout. Returns the character written, or WEOF on error.
Public Function fputwchar(ByVal c As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_fputwchar", False)
    Dim args(0 To 0) As Variant
    args(0) = c
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fputwchar = result
End Function

' FILE* _wfsopen(const wchar_t* filename, const wchar_t* mode, int shflag)
' Opens a file with a Unicode name and sharing flag. Returns FILE* or 0 on failure.
' shflag: SH_DENYRW=16, SH_DENYWR=32, SH_DENYRD=48, SH_DENYNO=64, SH_SECURE=128
Public Function wfsopen(ByVal filename As String, ByVal mode As String, _
                        ByVal shflag As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wfsopen", False)
    Dim args(0 To 2) As Variant
    args(0) = StrPtr(filename)
    args(1) = StrPtr(mode)
    args(2) = shflag
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wfsopen = result
End Function

' int fprintf(FILE* stream, const char* format [, ...])
' Writes formatted output to a stream (ANSI format).
' Format string is converted from VB6 Unicode to ANSI.
' NOTE: Variadic string args are not forwarded — compose the full string in VB first.
Public Function fprintf(ByVal stream As Long, ByVal format As String) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("fprintf")
    Dim formatBytes() As Byte
    formatBytes = StrConv(format, vbFromUnicode)
    Dim args(0 To 1) As Variant
    args(0) = stream
    args(1) = VarPtr(formatBytes(0))
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fprintf = result
End Function

' int fwprintf(FILE* stream, const wchar_t* format [, ...])
' Writes formatted wide output to a stream. Forwards variadic args directly.
Public Function fwprintf(ByVal stream As Long, ByVal format As String, _
                 ParamArray args() As Variant) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("fwprintf")
    
    Dim variadicCount As Long
    On Error Resume Next
    variadicCount = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then variadicCount = 0
    Err.Clear
    On Error GoTo 0
    
    Dim callArgs() As Variant
    ReDim callArgs(0 To 1 + variadicCount)
    callArgs(0) = stream
    callArgs(1) = StrPtr(format)
    
    Dim i As Long
    For i = 1 To variadicCount
        callArgs(1 + i) = args(LBound(args) + i - 1)
    Next
    
    Dim result As Variant
    fn.CallFunc vbLong, result, callArgs
    fwprintf = result
End Function

' size_t fwrite(const void* buffer, size_t size, size_t count, FILE* stream)
' Writes data from a buffer to a stream. Returns number of items written.
' buffer is passed as a raw Long pointer — use VarPtr() or StrPtr() to pass data.
Public Function fwrite(ByVal buffer As Long, ByVal size As Long, _
                       ByVal count As Long, ByVal stream As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("fwrite")
    Dim args(0 To 3) As Variant
    args(0) = buffer
    args(1) = size
    args(2) = count
    args(3) = stream
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    fwrite = result
End Function

' FILE* tmpfile(void)
' Creates a temporary file opened in w+b mode. Returns FILE* or 0 on failure.
' The file is automatically deleted when the stream is closed.
Public Function tmpfile() As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("tmpfile")
    Dim args(0 To 0) As Variant
    args(0) = 0&  ' dummy — tmpfile() takes no args
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    tmpfile = result
End Function
