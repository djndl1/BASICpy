Attribute VB_Name = "ReprImpl"
Option Explicit

' Core implementation for PyRepr — returns a string representation of a value,
' matching Python's repr() as closely as possible within VB6 constraints.
'
' Python semantics (approximated):
'   repr("hello")  → "'hello'"   (with quotes)
'   repr('ab"c')   → "'ab\"c'"   (single-quote outer, inner double-quote kept)
'   repr(42)       → '42'
'   repr(True)     → 'True'
'
' VB6: Differentiates VBA-specific types (Missing, Nothing, Empty, Null)
' even though Python does not have them.
'
' Return type is String (repr always returns a string).
Public Function PyReprCore(Optional ByVal value As Variant) As String
    ' No argument → 'Missing' (differentiates VBA's Missing from Python's None)
    If IsMissing(value) Then
        PyReprCore = "'Missing'"
        Exit Function
    End If

    Select Case VarType(value)
        Case vbEmpty
            PyReprCore = "'Empty'"

        Case vbNull
            PyReprCore = "'Null'"

        Case vbBoolean
            If value Then PyReprCore = "'True'" Else PyReprCore = "'False'"

        Case vbString
            Dim s As String: s = value

            Dim hasSingle As Boolean: hasSingle = (InStr(s, "'") > 0)
            Dim hasDouble As Boolean: hasDouble = (InStr(s, """") > 0)

            If hasSingle And Not hasDouble Then
                ' Contains single quotes but no double quotes — use double quotes
                ' as outer delimiter (Python behavior)
                PyReprCore = """" & s & """"
            ElseIf hasSingle And hasDouble Then
                ' Contains both — escape single quotes with backslash, use '
                ' as outer delimiter
                PyReprCore = "'" & Replace$(s, "'", "\'") & "'"
            Else
                ' Simple case — wrap in single quotes
                PyReprCore = "'" & s & "'"
            End If

        Case vbObject
            ' Objects implementing IRepr control their repr via Repr()
            If TypeOf value Is IRepr Then
                Dim reprImpl As IRepr: Set reprImpl = value
                PyReprCore = reprImpl.Repr()
            ElseIf value Is Nothing Then
                PyReprCore = "'Nothing'"
            Else
                PyReprCore = "'" & TypeName(value) & "'"
            End If

        Case vbByte, vbInteger, vbLong
            PyReprCore = CStr(value)

        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            PyReprCore = CStr(value)

        Case vbDate
            PyReprCore = CStr(value)

        Case Else
            On Error Resume Next
            PyReprCore = CStr(value)
            If Err.Number <> 0 Then
                PyReprCore = "'" & TypeName(value) & "'"
                On Error GoTo 0
            End If
    End Select
End Function
