VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDeclLibraryFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Wraps a CDECL function pointer obtained from NativeLibrary.GetProcAddress.
' Provides a Call method that delegates to DispCall.CDeclCall.
'
' Usage:
'   Dim fn As New CDeclLibraryFunction
'   fn.Init lib.GetProcAddress("MyFunction")
'   Dim result As Variant
'   fn.CallFunc vbLong, result, Array(StrPtr(sOut), StrPtr(sFormat), lParam1)
'   ' Or use the pointer directly with CDeclCall:
'   ' DispCall.CDeclCall fn.Ptr, vbLong, result, arg1, arg2

Private m_fnPtr As Long

' Initialises the wrapper with a function pointer from NativeLibrary.GetProcAddress.
Public Sub Init(ByVal fnPtr As Long)
    m_fnPtr = fnPtr
End Sub

' Calls the wrapped function with the given return type and arguments.
'   resultType - VARTYPE of the return value (vbLong, vbString, vbEmpty for void, etc.)
'   resultVar  - Variant that receives the return value
'   args       - Variant array of positional arguments
'
' Returns the HRESULT from DispCallFunc (0 = S_OK).
' The actual function return value is written into resultVar.
Public Function CallFunc(ByVal resultType As Integer, _
                         ByRef resultVar As Variant, _
                         ByRef args() As Variant) As Long
    If m_fnPtr = 0 Then
        Err.Raise vbObjectError + 1027, "CDeclLibraryFunction.CallFunc", _
            "Function pointer is not initialised"
    End If
    
    CallFunc = DispCall.CDeclCall(m_fnPtr, resultType, resultVar, args)
End Function

' Returns the raw function pointer for direct use with DispCall.CDeclCall.
Public Property Get Ptr() As Long
    Ptr = m_fnPtr
End Property
