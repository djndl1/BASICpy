VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PyMath"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Returns the ceiling of a number, the smallest integer greater than or equal
' to it.  Calls CDbl on the input, then applies -Int(-x).
' Args:
'   value (Variant):  Numeric value to compute the ceiling of.
' Returns:
'   Double:  The ceiling of value.
Public Function PyCeil(ByVal value As Variant) As Double
Attribute PyCeil.VB_Description = "Returns the ceiling of a number, the smallest integer greater than or equal to it.\r\nArgs:\r\n  value (Variant): Numeric value to compute the ceiling of.\r\nReturns:\r\n  Double: The ceiling of value."
    PyCeil = -Int(-CDbl(value))
End Function

' Returns the floor of a number, the largest integer less than or equal to
' it.  Calls CDbl on the input, then applies Int (which rounds toward
' negative infinity in VB6).
' Args:
'   value (Variant):  Numeric value to compute the floor of.
' Returns:
'   Double:  The floor of value.
Public Function PyFloor(ByVal value As Variant) As Double
Attribute PyFloor.VB_Description = "Returns the floor of a number, the largest integer less than or equal to it.\r\nArgs:\r\n  value (Variant): Numeric value to compute the floor of.\r\nReturns:\r\n  Double: The floor of value."
    PyFloor = Int(CDbl(value))
End Function

' Returns the absolute value of x as a Double.  Wraps VB6's built-in Abs.
' Args:
'   value (Variant):  Numeric value to return the absolute of.
' Returns:
'   Double:  The absolute value of value.
Public Function PyFabs(ByVal value As Variant) As Double
Attribute PyFabs.VB_Description = "Returns the absolute value of x as a Double.\r\nArgs:\r\n  value (Variant): Numeric value to return the absolute of.\r\nReturns:\r\n  Double: The absolute value of value."
    PyFabs = Abs(CDbl(value))
End Function

' Returns x with the fractional part removed, rounding toward zero.
' Equivalent to floor() for positive x, ceil() for negative x.
' Uses VB6's Fix function which truncates toward zero.
' Args:
'   value (Variant):  Numeric value to truncate.
' Returns:
'   Double:  The truncated value.
Public Function PyTrunc(ByVal value As Variant) As Double
Attribute PyTrunc.VB_Description = "Returns x with the fractional part removed, rounding toward zero.\r\nArgs:\r\n  value (Variant): Numeric value to truncate.\r\nReturns:\r\n  Double: The truncated value."
    PyTrunc = Fix(CDbl(value))
End Function