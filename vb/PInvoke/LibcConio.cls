VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcConio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ─── Console input ───

' char* _cgets(char* buffer)
' Reads a line of input from the console (ANSI version).
' buffer[0] = maximum chars to read (excluding null) — set by caller
' buffer[1] = chars actually read — set by function
' buffer[2..] = null-terminated ANSI string data
' Returns pointer to buffer[2] (the start of the string).
Public Function cgets(ByRef buffer() As Byte) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_cgets", False)
    Dim args(0 To 0) As Variant
    args(0) = VarPtr(buffer(0))
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    cgets = result
End Function

' wchar_t* _cgetws(wchar_t* buffer)
' Reads a line of input from the console (Unicode version).
' buffer(0..1) = maximum wchar_t to read (first wchar_t) — set by caller
' buffer(2..3) = wchar_t actually read (second wchar_t) — set by function
' buffer(4..) = null-terminated wide string data
' Returns pointer to buffer(4) (the start of the string).
Public Function cgetws(ByRef buffer() As Byte) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_cgetws", False)
    Dim args(0 To 0) As Variant
    args(0) = VarPtr(buffer(0))
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    cgetws = result
End Function

' int _cputws(const wchar_t* str)
' Writes a wide string directly to the console. Returns 0 on success, EOF on failure.
Public Function cputws(ByVal str As String) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("cputws")
    Dim args(0 To 0) As Variant
    args(0) = StrPtr(str)
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    cputws = result
End Function

' int _cscanf(const char* format [, argument] ...)
' Reads formatted input from the console (ANSI version).
' NOTE: Variadic args are OUTPUT pointers — not forwardable through this wrapper.
'       Pass the format string only. For actual scanf use, wrap via a custom helper.
Public Function cscanf(ByVal format As String) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("cscanf")
    Dim formatBytes() As Byte
    formatBytes = StrConv(format, vbFromUnicode)
    Dim args(0 To 0) As Variant
    args(0) = VarPtr(formatBytes(0))
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    cscanf = result
End Function

' ─── Console output (printf-style) ───

' int _cwprintf(const wchar* format [, argument_list])
' Writes formatted wide output to the console. Forwards variadic args directly
' to the CDECL call. Args must match format specifiers in type:
'   String → %s (StrPtr gives wchar_t* — BSTR data IS null-terminated wide chars)
'   Long   → %d
'   Double → %f
'   Integer → %hd
Public Function cwprintf(ByVal format As String, ParamArray args() As Variant) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("cwprintf")
    
    Dim variadicCount As Long
    On Error Resume Next
    variadicCount = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then variadicCount = 0
    Err.Clear
    On Error GoTo 0
    
    Dim callArgs() As Variant
    ReDim callArgs(0 To variadicCount)
    callArgs(0) = StrPtr(format)
    
    Dim i As Long
    For i = 1 To variadicCount
        callArgs(i) = args(LBound(args) + i - 1)
    Next
    
    Dim result As Variant
    fn.CallFunc vbLong, result, callArgs
    cwprintf = result
End Function

' int _cwprintf_l(const wchar* format, _locale_t locale [, argument_list])
' Writes formatted wide output to the console using the specified locale.
Public Function cwprintf_l(ByVal format As String, _
                           ByVal locale As Long, _
                   ParamArray args() As Variant) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("cwprintf_l")
    
    Dim variadicCount As Long
    On Error Resume Next
    variadicCount = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then variadicCount = 0
    Err.Clear
    On Error GoTo 0
    
    Dim callArgs() As Variant
    ReDim callArgs(0 To 1 + variadicCount)
    callArgs(0) = StrPtr(format)
    callArgs(1) = locale
    
    Dim i As Long
    For i = 1 To variadicCount
        callArgs(1 + i) = args(LBound(args) + i - 1)
    Next
    
    Dim result As Variant
    fn.CallFunc vbLong, result, callArgs
    cwprintf_l = result
End Function
