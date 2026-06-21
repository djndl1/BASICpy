Attribute VB_Name = "AbsImpl"
Option Explicit

Public Function PyAbsCore(ByVal value As Variant) As Variant
    If IsObject(value) Then
        If TypeOf value Is IAbsolute Then
            Dim absImpl As IAbsolute
            Set absImpl = value
            PyAbs = absImpl.Absolute()
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyAbs", "object does not implement IAbsolute"
        End If
    Else
        PyAbs = Abs(value)
    End If
End Function
