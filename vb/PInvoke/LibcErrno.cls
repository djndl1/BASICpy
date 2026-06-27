VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcErrno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef dest As Any, ByRef src As Any, ByVal length As Long)

Private m_fn_errno As CDeclLibraryFunction

Public Enum ErrnoValue
    EPERM = 1
    ENOENT = 2
    ESRCH = 3
    EINTR = 4
    EIO = 5
    ENXIO = 6
    E2BIG = 7
    ENOEXEC = 8
    EBADF = 9
    ECHILD = 10
    EAGAIN = 11
    ENOMEM = 12
    EACCES = 13
    EFAULT = 14
    EBUSY = 16
    EEXIST = 17
    EXDEV = 18
    ENODEV = 19
    ENOTDIR = 20
    EISDIR = 21
    EINVAL = 22
    ENFILE = 23
    EMFILE = 24
    ENOTTY = 25
    EFBIG = 27
    ENOSPC = 28
    ESPIPE = 29
    EROFS = 30
    EMLINK = 31
    EPIPE = 32
    EDOM = 33
    ERANGE = 34
    EDEADLK = 36
    ENAMETOOLONG = 38
    ENOLCK = 39
    ENOSYS = 40
    ENOTEMPTY = 41
    EILSEQ = 42
    STRUNCATE = 80
    EADDRINUSE = 100
    EADDRNOTAVAIL = 101
    EAFNOSUPPORT = 102
    EALREADY = 103
    EBADMSG = 104
    ECANCELED = 105
    ECONNABORTED = 106
    ECONNREFUSED = 107
    ECONNRESET = 108
    EDESTADDRREQ = 109
    EHOSTUNREACH = 110
    EIDRM = 111
    EINPROGRESS = 112
    EISCONN = 113
    ELOOP = 114
    EMSGSIZE = 115
    ENETDOWN = 116
    ENETRESET = 117
    ENETUNREACH = 118
    ENOBUFS = 119
    ENODATA = 120
    ENOLINK = 121
    ENOMSG = 122
    ENOPROTOOPT = 123
    ENOSR = 124
    ENOSTR = 125
    ENOTCONN = 126
    ENOTRECOVERABLE = 127
    ENOTSOCK = 128
    ENOTSUP = 129
    EOPNOTSUPP = 130
    EOVERFLOW = 132
    EOWNERDEAD = 133
    EPROTO = 134
    EPROTONOSUPPORT = 135
    EPROTOTYPE = 136
    ETIME = 137
    ETIMEDOUT = 138
    ETXTBSY = 139
    EWOULDBLOCK = 140
End Enum

Private Sub Class_Initialize()
    Set m_fn_errno = LibcLibrary.ResolveFunction("_errno")
End Sub

Private Sub Class_Terminate()
    Set m_fn_errno = Nothing
End Sub

' int* _errno(void)
' Returns a pointer to the thread-local errno value.
Private Function ErrnoPtr() As Long
    Dim args(0 To 0) As Variant
    args(0) = 0&  ' dummy — _errno() takes no args, ignores stack
    Dim result As Variant
    m_fn_errno.CallFunc vbLong, result, args
    ErrnoPtr = CLng(result)
End Function

' Reads the current thread-local errno value via _errno() pointer dereference.
Public Property Get Current() As Long
    Dim val As Long
    Dim ptr As Long: ptr = ErrnoPtr
    If ptr <> 0 Then CopyMemory val, ByVal ptr, 4
    Current = val
End Property

' Sets the thread-local errno value via _errno() pointer dereference.
Public Sub SetValue(ByVal value As ErrnoValue)
    Dim ptr As Long: ptr = ErrnoPtr
    If ptr <> 0 Then CopyMemory ByVal ptr, value, 4
End Sub

' Sets errno to 0.
Public Sub Clear()
    SetValue 0
End Sub

' wchar_t* _wcserror(int errnum)
' Returns a human-readable description for the given errno value.
' The returned string comes from a C runtime static buffer — it is copied
' out immediately. Use LibcErrno.Describe(code) in place of strerror().
Public Function wcserror(ByVal errnum As Long) As String
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_wcserror", False)
    Dim args(0 To 0) As Variant
    args(0) = errnum
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    wcserror = PtrToWStr(CLng(result))
End Function

' Copies a C wchar_t* (null-terminated wide string) into a VB String.
Private Function PtrToWStr(ByVal ptr As Long) As String
    If ptr = 0 Then Exit Function
    Dim chars As Long
    chars = lstrlenW(ptr)
    If chars = 0 Then Exit Function
    Dim buf() As Byte
    ReDim buf(0 To chars * 2 - 1) As Byte
    CopyMemory buf(0), ByVal ptr, chars * 2
    PtrToWStr = buf
End Function
