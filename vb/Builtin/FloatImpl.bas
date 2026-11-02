Attribute VB_Name = "FloatImpl"
Option Explicit

' Core implementation for PyFloat — converts a value to a floating-point number,
' matching Python's float().
'
' Python semantics:
'   float()              → 0.0
'   float(42)            → 42.0
'   float(3.14)          → 3.14
'   float(True)          → 1.0
'   float(False)         → 0.0
'   float("3.14")        → 3.14
'   float(obj)           → obj.__float__() if defined, else obj.__index__()
'
' VB6: Returns Double (VB6's native floating-point type, matching Python's float).
'      Objects implementing IFloat are converted via Float().
'      Objects implementing IIndex (but not IFloat) fall back to Index().
Public Function PyFloatCore(Optional ByVal value As Variant) As Double
    ' No argument → 0.0 (Python float() → 0.0)
    If IsMissing(value) Then
        PyFloatCore = 0#
        Exit Function
    End If

    Select Case VarType(value)
        Case vbEmpty
            PyFloatCore = 0#

        Case vbNull
            Err.Raise CommonHResults.TypeMismatch, "Builtin.PyFloat", _
                "float() argument must be a string or a number, not 'NoneType'"

        Case vbBoolean
            If value Then
                PyFloatCore = 1#
            Else
                PyFloatCore = 0#
            End If

        Case vbByte, vbInteger, vbLong
            PyFloatCore = CDbl(value)

        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            PyFloatCore = CDbl(value)

        Case vbString
            ' Parse string as float (VB6's CDbl handles leading/trailing whitespace)
            Dim s As String: s = Trim$(value)
            Ensure.IsTrue Len(s) > 0, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyFloat", "could not convert string to float: empty string"
            On Error Resume Next
            PyFloatCore = CDbl(s)
            If Err.Number <> 0 Then
                Err.Raise CommonHResults.InvalidProcedureCall, "Builtin.PyFloat", _
                    "could not convert string to float: '" & s & "'"
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0

        Case vbObject
            ' Python's float(x) delegates to x.__float__() if defined,
            ' otherwise falls back to x.__index__().
            If TypeOf value Is IFloat Then
                Dim floatImpl As IFloat: Set floatImpl = value
                PyFloatCore = floatImpl.Float()
            ElseIf TypeOf value Is IIndex Then
                Dim idx As IIndex: Set idx = value
                PyFloatCore = CDbl(Builtin.PyInt(value))
            Else
                Err.Raise CommonHResults.TypeMismatch, "Builtin.PyFloat", _
                    "float() argument must be a string or a number, not '" & TypeName(value) & "'"
            End If

        Case Else
            Err.Raise CommonHResults.TypeMismatch, "Builtin.PyFloat", _
                "float() argument must be a string or a number"
    End Select
End Function
