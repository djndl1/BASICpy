Attribute VB_Name = "RangeImpl"
Option Explicit

' Core implementation — called by Builtin.PyRange facade.
Public Function PyRangeCore(Optional ByVal arg1 As Variant, _
                        Optional ByVal arg2 As Variant, _
                        Optional ByVal arg3 As Long = 1) As RangeIterator
    Dim startVal As Long
    Dim stopVal As Long
    Dim stepVal As Long: stepVal = arg3
    
    If IsMissing(arg1) Then
        startVal = 0: stopVal = 0: stepVal = 1
    ElseIf IsMissing(arg2) Then
        startVal = 0
        stopVal = AsIntegerIndex(arg1)
        stepVal = 1
    Else
        startVal = AsIntegerIndex(arg1)
        stopVal = AsIntegerIndex(arg2)
    End If
    
    Dim r As New RangeIterator
    r.Init startVal, stopVal, stepVal
    Set PyRangeCore = r
End Function
