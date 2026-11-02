Attribute VB_Name = "TypeUtil"
Option Explicit

' Shared type-checking utilities for BASICpy builtins.

' Returns True if the Variant holds an integer value.
'
' Integer types: vbByte, vbInteger, vbLong.
' Floating-point types: vbSingle, vbDouble, vbCurrency, vbDecimal —
'   considered integer only if they hold a whole number (Fix(v) = v).
Public Function IsIntegerValue(ByVal v As Variant) As Boolean
    If IsObject(v) Then
        IsIntegerValue = False
        Exit Function
    End If

    Select Case VarType(v)
        Case vbByte, vbInteger, vbLong
            IsIntegerValue = True
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            IsIntegerValue = (Fix(v) = v)
        Case Else
            IsIntegerValue = False
    End Select
End Function
