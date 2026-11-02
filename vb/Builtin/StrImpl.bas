Attribute VB_Name = "StrImpl"
Option Explicit

' Core implementation for PyStr — converts a value to its string representation,
' matching Python's str() as closely as possible within VB6 constraints.
'
' Python semantics approximated:
'   str()          → ''
'   str(None)      → 'None'
'   str(True)      → 'True'
'   str(False)     → 'False'
'   str(42)        → '42'
'   str(3.14)      → '3.14'
'   str("hello")   → 'hello'   (strings returned as-is)
'   str(object)    → TypeName-based representation
Public Function PyStrCore(Optional ByVal value As Variant) As Variant
    ' No argument → empty string (Python str() → '')
    If IsMissing(value) Then
        PyStrCore = ""
        Exit Function
    End If

    ' Nothing → 'None'
    If IsObject(value) Then
        If value Is Nothing Then
            PyStrCore = "None"
            Exit Function
        End If
    End If

    Select Case VarType(value)
        Case vbString
            ' Strings returned as-is
            PyStrCore = value

        Case vbBoolean
            If value Then
                PyStrCore = "True"
            Else
                PyStrCore = "False"
            End If

        Case vbEmpty
            PyStrCore = ""

        Case vbNull
            PyStrCore = "None"

        Case vbByte, vbInteger, vbLong
            PyStrCore = CStr(value)

        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            PyStrCore = CStr(value)

        Case vbDate
            PyStrCore = CStr(value)

        Case vbObject
            ' Objects: try to get a meaningful string representation
            If TypeOf value Is IBoolean Then
                Dim boolImpl As IBoolean: Set boolImpl = value
                If boolImpl.Bool() Then
                    PyStrCore = "True"
                Else
                    PyStrCore = "False"
                End If
            Else
                PyStrCore = TypeName(value)
            End If

        Case Else
            ' Fallback: use CStr
            On Error Resume Next
            PyStrCore = CStr(value)
            If Err.Number <> 0 Then
                PyStrCore = TypeName(value)
                On Error GoTo 0
            End If
    End Select
End Function
