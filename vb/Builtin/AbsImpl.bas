Attribute VB_Name = "AbsImpl"
Option Explicit

Public Function PyAbsCore(ByVal value As Variant) As Variant
    ' Fast path: numeric scalars
    If Not IsObject(value) Then
        PyAbsCore = Abs(value)
        Exit Function
    End If

    ' Guard: objects must implement IAbsolute
    Ensure.IsTrue TypeOf value Is IAbsolute, CommonHResults.InvalidProcedureCall, _
        "Builtin.PyAbs", "object does not implement IAbsolute"

    ' Object implementing IAbsolute
    Dim absImpl As IAbsolute: Set absImpl = value
    PyAbsCore = absImpl.Absolute()
End Function