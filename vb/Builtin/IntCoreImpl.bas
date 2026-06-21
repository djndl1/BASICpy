Attribute VB_Name = "IntCoreImpl"
Option Explicit

' Core implementation for PyInt — returns an integer matching Python's int().
Public Function PyIntCore(Optional ByVal value As Variant, Optional ByVal base As Variant) As Variant
    If IsArray(value) Then
        Ensure.IsTrue False, CommonHResults.TypeMismatch, _
            "Builtin.PyInt", "argument must be a number or string, not an array"
    
    ElseIf IsMissing(value) Then
        PyIntCore = 0
    
    ElseIf Not IsMissing(base) Then
        PyIntCore = ParseIntString(CStr(value), AsIntegerIndex(base))
    
    ElseIf IsObject(value) Then
        If TypeOf value Is IIndex Then
            Dim idx As IIndex
            Set idx = value
            PyIntCore = PyInt(idx.index())
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyInt", "object does not implement IIndex"
        End If
    
    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong
                PyIntCore = value
            Case vbSingle, vbDouble, vbCurrency, vbDecimal
                PyIntCore = Fix(value)
            Case vbString
                PyIntCore = ParseIntString(value, 10)
            Case Else
                Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                    "Builtin.PyInt", "argument must be a number or string"
        End Select
    End If
End Function
