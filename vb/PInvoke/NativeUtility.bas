Attribute VB_Name = "NativeUtility"
Option Explicit

' Copies a C wchar_t* (null-terminated wide string) into a VB String.
Public Function NullTerminatedWStringToBStr(ByVal ptr As Long) As String
    If ptr = 0 Then Exit Function

    Dim chars As Long: chars = lstrlenW(ptr)
    If chars = 0 Then Exit Function

    Dim buf As String: buf = Space$(chars)
    lstrcpyW StrPtr(buf), ptr
    NullTerminatedWStringToBStr = buf
End Function

' Copies a VB String to a null-terminated wide string buffer.
Public Sub CopyStringToNullTerminatedWString(ByVal dst As Long, ByRef src As String)
    If dst = 0 Then Exit Sub
    lstrcpyW dst, StrPtr(src)
End Sub