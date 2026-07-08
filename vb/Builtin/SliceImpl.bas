Attribute VB_Name = "SliceImpl"
Option Explicit

' ─── PySliceCore ───
' Core implementation — called by Builtin.PySlice facade.
' Handles Python's varargs:
'   slice(stop)
'   slice(start, stop[, step])
'
' Missing args become Empty (our representation of Python's None).

Public Function PySliceCore(Optional ByVal arg1 As Variant, _
                            Optional ByVal arg2 As Variant, _
                            Optional ByVal arg3 As Variant) As Slice

    Dim start As Variant
    Dim stopVal As Variant
    Dim stepVal As Variant

    If IsMissing(arg1) Then
        ' slice() → slice(None, None, None)
        start = Empty
        stopVal = Empty
        stepVal = Empty
    ElseIf IsMissing(arg2) Then
        ' slice(stop) → slice(None, stop, None)
        start = Empty
        stopVal = arg1
        stepVal = Empty
    Else
        ' slice(start, stop[, step])
        start = arg1
        stopVal = arg2
        If IsMissing(arg3) Then
            stepVal = Empty
        Else
            stepVal = arg3
        End If
    End If

    Dim s As New Slice
    s.Init start, stopVal, stepVal
    Set PySliceCore = s
End Function
