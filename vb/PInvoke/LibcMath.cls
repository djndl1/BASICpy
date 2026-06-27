VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' All functions resolve through LibcLibrary.ResolveFunction() which caches and
' handles CDECL name decoration. Functions with underscore in the
' name (e.g., _cabs) use mangle=False.

' double fabs(double x)
Public Function fabs(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("fabs")
    Dim args(0 To 0) As Variant
    args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    fabs = result
End Function

' struct _complex { double x; double y; }
' double _cabs(struct _complex z)
' On x86 Win32 CDECL, a struct {double,double} by value is ABI-equivalent
' to two separate double parameters, so we pass x and y directly.
Public Function cabs(ByVal x As Double, ByVal y As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_cabs", False)
    Dim args(0 To 1) As Variant
    args(0) = x
    args(1) = y
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    cabs = result
End Function

' long labs(long n)
Public Function labs(ByVal n As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("labs")
    Dim args(0 To 0) As Variant
    args(0) = n
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    labs = result
End Function

' double acos(double x)
Public Function acos(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("acos")
    Dim args(0 To 0) As Variant
    args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    acos = result
End Function

' double asin(double x)
Public Function asin(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("asin")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    asin = result
End Function

' double atan(double x)
Public Function atan(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("atan")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    atan = result
End Function

' double atan2(double y, double x)
Public Function atan2(ByVal y As Double, ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("atan2")
    Dim args(0 To 1) As Variant: args(0) = y: args(1) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    atan2 = result
End Function

' double cosh(double x)
Public Function cosh(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("cosh")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    cosh = result
End Function

' double tanh(double x)
Public Function tanh(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("tanh")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    tanh = result
End Function

' double tan(double x)
Public Function tan(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("tan")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    tan = result
End Function

' double log(double x)
Public Function log(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("log")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    log = result
End Function

' double log10(double x)
Public Function log10(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("log10")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    log10 = result
End Function

' double ceil(double x)
Public Function ceil(ByVal x As Double) As Double
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("ceil")
    Dim args(0 To 0) As Variant: args(0) = x
    Dim result As Variant
    fn.CallFunc vbDouble, result, args
    ceil = result
End Function
