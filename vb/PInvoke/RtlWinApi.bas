Attribute VB_Name = "RtlWinApi"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByVal destination As Long, ByVal source As Long, ByVal length As Long)

Public Declare Sub FillMemory Lib "kernel32" (destination As Any, ByVal fill As Byte, ByVal length As Long)
' Memory fill operation (similar to memset)
Public Declare Sub RtlFillMemory Lib "kernel32" (ByVal destination As Long, ByVal fill As Byte, ByVal length As Long)

Public Declare Function CompareMemory Lib "kernel32" (source1 As Any, ByVal source2 As Long, ByVal length As Long) As Long
' Memory comparison operation (similar to memcmp)
Public Declare Function RtlCompareMemory Lib "kernel32" (ByVal source1 As Long, ByVal source2 As Long, ByVal length As Long) As Long

Public Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Public Declare Function CoTaskMemAlloc Lib "Ole32" (ByVal cb As Long) As Long

Public Declare Sub CoTaskMemFree Lib "Ole32" (ByVal pv As Long)

Public Declare Function CoTaskMemRealloc Lib "Ole32" (ByVal pv As Long, ByVal cb As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long
