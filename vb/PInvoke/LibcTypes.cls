VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' C type aliases (for documentation / self-documenting code)
' size_t       → Long
' intptr_t     → Long
' FILE*        → Long
' wchar_t*     → Long (StrPtr)
' errno_t      → Long

' ─── stdio.h ───

Public Enum StdioConsts
    EOF = -1
    SEEK_SET = 0
    SEEK_CUR = 1
    SEEK_END = 2
End Enum

' ─── fcntl.h — file open flags ───

Public Enum FcntlFlags
    O_RDONLY = &H0
    O_WRONLY = &H1
    O_RDWR = &H2
    O_ACCMODE = &H2       ' O_RDONLY Or O_WRONLY Or O_RDWR
    O_APPEND = &H8
    O_RANDOM = &H10
    O_SEQUENTIAL = &H20
    O_TEMPORARY = &H40
    O_NOINHERIT = &H80
    O_CREAT = &H100
    O_TRUNC = &H200
    O_EXCL = &H400
    O_SHORT_LIVED = &H1000
    O_TEXT = &H4000
    O_BINARY = &H8000
    O_RAW = &H8000        ' alias for O_BINARY
    O_WTEXT = &H10000
    O_U16TEXT = &H20000
    O_U8TEXT = &H40000
End Enum

' ─── sys/stat.h — file permission bits ───

Public Enum StatFlags
    S_IWRITE = &H80
    S_IREAD = &H100
End Enum

' ─── io.h — standard file descriptors ───

Public Enum IoConsts
    STDIN_FILENO = 0
    STDOUT_FILENO = 1
    STDERR_FILENO = 2
End Enum

' ─── stdlib.h ───

Public Enum StdlibConsts
    EXIT_SUCCESS = 0
    EXIT_FAILURE = 1
End Enum
