VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ensure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Ensures that an object reference is not Nothing
' Args:
'   obj (Object): The object reference to check
'   source (String): Optional source identifier for error reporting
Public Sub NotNothing(ByVal obj As Object, Optional source As String = "", Optional message As String = "")
Attribute NotNothing.VB_Description = "Ensures that an object reference is not Nothing\r\nArgs:\r\n  obj (Object): The object reference to check\r\n  source (String): Optional source identifier for error reporting\r\n  message (String): Optional custom error message"
    If message = "" Then
        IsTrue (obj Is Nothing), CommonHResults.ObjectVariableNotSet, source
    Else
        IsTrue (obj Is Nothing), CommonHResults.ObjectVariableNotSet, source, message
    End If
End Sub

' Ensures that a variant has the specified type
' Args:
'   v (Variant): The variant to check
'   t (VbVarType): The expected variable type
'   source (String): Optional source identifier for error reporting
Public Sub VariantType(ByRef v As Variant, ByVal t As VbVarType, Optional source As String = "", Optional message As String = "")
Attribute VariantType.VB_Description = "Ensures that a variant has the specified type\r\nArgs:\r\n  v (Variant): The variant to check\r\n  t (VbVarType): The expected variable type\r\n  source (String): Optional source identifier for error reporting\r\n  message (String): Optional custom error message"
    If message = "" Then
        IsTrue VarType(v) = t, CommonHResults.TypeMismatch, source
    Else
        IsTrue VarType(v) = t, CommonHResults.TypeMismatch, source, message
    End If
End Sub

' Raises an error if the condition is true
' Args:
'   b (Boolean): The condition to check
'   errCode (Long): The error code to raise if condition is true
'   source (String): Optional source identifier for error reporting
Public Sub IsFalse(ByVal b As Boolean, ByVal errCode As Long, Optional source As String = "", Optional message As String = "")
Attribute IsFalse.VB_Description = "Raises an error if the condition is true\r\nArgs:\r\n  b (Boolean): The condition to check\r\n  errCode (Long): The error code to raise if condition is true\r\n  source (String): Optional source identifier for error reporting\r\n  message (String): Optional custom error message"
    If b Then
        If message = "" Then
            Err.Raise errCode, source
        Else
            Err.Raise errCode, source, message
        End If
    End If
End Sub

' Raises an error if the condition is false
' Args:
'   b (Boolean): The condition to check
'   errCode (Long): The error code to raise if condition is false
'   source (String): Optional source identifier for error reporting
Public Sub IsTrue(ByVal b As Boolean, ByVal errCode As Long, Optional source As String = "", Optional message As String = "")
Attribute IsTrue.VB_Description = "Raises an error if the condition is false\r\nArgs:\r\n  b (Boolean): The condition to check\r\n  errCode (Long): The error code to raise if condition is false\r\n  source (String): Optional source identifier for error reporting\r\n  message (String): Optional custom error message"
    If Not b Then
        If message = "" Then
            Err.Raise errCode, source
        Else
            Err.Raise errCode, source, message
        End If
    End If
End Sub

' Ensures that two values are not equal
' Args:
'   a (Variant): First value to compare
'   b (Variant): Second value to compare
'   errCode (Long): The error code to raise if values are equal
'   source (String): Optional source identifier for error reporting
Public Sub NotEqual(ByVal a As Variant, ByVal b As Variant, ByVal errCode As Long, Optional source As String = "", Optional message As String = "")
Attribute NotEqual.VB_Description = "Ensures that two values are not equal\r\nArgs:\r\n  a (Variant): First value to compare\r\n  b (Variant): Second value to compare\r\n  errCode (Long): The error code to raise if values are equal\r\n  source (String): Optional source identifier for error reporting\r\n  message (String): Optional custom error message"
    IsFalse a = b, errCode, source, message
End Sub

' Ensures that two values are equal
' Args:
'   a (Variant): First value to compare
'   b (Variant): Second value to compare
'   errCode (Long): The error code to raise if values are not equal
'   source (String): Optional source identifier for error reporting
Public Sub AreEqual(ByVal a As Variant, ByVal b As Variant, ByVal errCode As Long, Optional source As String = "", Optional message As String = "")
Attribute AreEqual.VB_Description = "Ensures that two values are equal\r\nArgs:\r\n  a (Variant): First value to compare\r\n  b (Variant): Second value to compare\r\n  errCode (Long): The error code to raise if values are not equal\r\n  source (String): Optional source identifier for error reporting\r\n  message (String): Optional custom error message"
    IsTrue a = b, errCode, source, message
End Sub

' Guards-style parameter validation methods

' Ensures that a value is greater than a reference value
' Args:
'   value (Variant): The value to check
'   refValue (Variant): The reference value to compare against
'   paramName (String): The parameter name for error reporting
Public Sub GreaterThan(ByVal value As Variant, ByVal refValue As Variant, ByVal paramName As String)
Attribute GreaterThan.VB_Description = "Ensures that a value is greater than a reference value\r\nArgs:\r\n  value (Variant): The value to check\r\n  refValue (Variant): The reference value to compare against\r\n  paramName (String): The parameter name for error reporting"
    If value <= refValue Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.GreaterThan", paramName & " cannot be less than or equal to " & CStr(refValue)
    End If
End Sub

' Ensures that a value is greater than or equal to a reference value
' Args:
'   value (Variant): The value to check
'   refValue (Variant): The reference value to compare against
'   paramName (String): The parameter name for error reporting
Public Sub GreaterThanOrEqual(ByVal value As Variant, ByVal refValue As Variant, ByVal paramName As String)
Attribute GreaterThanOrEqual.VB_Description = "Ensures that a value is greater than or equal to a reference value\r\nArgs:\r\n  value (Variant): The value to check\r\n  refValue (Variant): The reference value to compare against\r\n  paramName (String): The parameter name for error reporting"
    If value < refValue Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.GreaterThanOrEqual", paramName & " cannot be less than " & CStr(refValue)
    End If
End Sub

' Ensures that a value is less than a reference value
' Args:
'   value (Variant): The value to check
'   refValue (Variant): The reference value to compare against
'   paramName (String): The parameter name for error reporting
Public Sub LessThan(ByVal value As Variant, ByVal refValue As Variant, ByVal paramName As String)
Attribute LessThan.VB_Description = "Ensures that a value is less than a reference value\r\nArgs:\r\n  value (Variant): The value to check\r\n  refValue (Variant): The reference value to compare against\r\n  paramName (String): The parameter name for error reporting"
    If value >= refValue Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.LessThan", paramName & " cannot be greater than or equal to " & CStr(refValue)
    End If
End Sub

' Ensures that a value is less than or equal to a reference value
' Args:
'   value (Variant): The value to check
'   refValue (Variant): The reference value to compare against
'   paramName (String): The parameter name for error reporting
Public Sub LessThanOrEqual(ByVal value As Variant, ByVal refValue As Variant, ByVal paramName As String)
Attribute LessThanOrEqual.VB_Description = "Ensures that a value is less than or equal to a reference value\r\nArgs:\r\n  value (Variant): The value to check\r\n  refValue (Variant): The reference value to compare against\r\n  paramName (String): The parameter name for error reporting"
    If value > refValue Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.LessThanOrEqual", paramName & " cannot be greater than " & CStr(refValue)
    End If
End Sub

' Ensures that a value is between minimum and maximum values
' Args:
'   value (Variant): The value to check
'   min (Variant): The minimum allowed value
'   max (Variant): The maximum allowed value
'   inclusive (Boolean): Whether to include boundaries
'   paramName (String): The parameter name for error reporting
Public Sub Between(ByVal value As Variant, ByVal min As Variant, ByVal max As Variant, ByVal inclusive As Boolean, ByVal paramName As String)
Attribute Between.VB_Description = "Ensures that a value is between minimum and maximum values\r\nArgs:\r\n  value (Variant): The value to check\r\n  min (Variant): The minimum allowed value\r\n  max (Variant): The maximum allowed value\r\n  inclusive (Boolean): Whether to include boundaries\r\n  paramName (String): The parameter name for error reporting"
    Dim atBoundary As Boolean
    atBoundary = (value = min) Or (value = max)
    
    If value < min Or value > max Or (Not inclusive And atBoundary) Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.Between", paramName & " is out of the range between " & CStr(min) & " and " & CStr(max) & "!"
    End If
End Sub

' Ensures that a boolean value is not True
' Args:
'   b (Boolean): The boolean value to check
'   paramName (String): The parameter name for error reporting
Public Sub NotTrue(ByVal b As Boolean, ByVal paramName As String)
Attribute NotTrue.VB_Description = "Ensures that a boolean value is not True\r\nArgs:\r\n  b (Boolean): The boolean value to check\r\n  paramName (String): The parameter name for error reporting"
    If b Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.NotTrue", paramName & " cannot be true!"
    End If
End Sub

' Ensures that a boolean value is not False
' Args:
'   b (Boolean): The boolean value to check
'   paramName (String): The parameter name for error reporting
Public Sub NotFalse(ByVal b As Boolean, ByVal paramName As String)
Attribute NotFalse.VB_Description = "Ensures that a boolean value is not False\r\nArgs:\r\n  b (Boolean): The boolean value to check\r\n  paramName (String): The parameter name for error reporting"
    If Not b Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.NotFalse", paramName & " cannot be false!"
    End If
End Sub

' Ensures that an object reference is not Nothing
' Args:
'   obj (Object): The object reference to check
'   paramName (String): The parameter name for error reporting
Public Sub ArgumentNotNull(ByVal obj As Object, ByVal paramName As String)
Attribute ArgumentNotNull.VB_Description = "Ensures that an object reference is not Nothing\r\nArgs:\r\n  obj (Object): The object reference to check\r\n  paramName (String): The parameter name for error reporting"
    If obj Is Nothing Then
        Err.Raise CommonHResults.ObjectVariableNotSet, "Ensure.ArgumentNotNull", paramName & " cannot be Nothing!"
    End If
End Sub

' Ensures that a string is not empty or null
' Args:
'   strValue (String): The string value to check
'   paramName (String): The parameter name for error reporting
Public Sub StringNotEmpty(ByVal strValue As String, ByVal paramName As String)
Attribute StringNotEmpty.VB_Description = "Ensures that a string is not empty or null\r\nArgs:\r\n  strValue (String): The string value to check\r\n  paramName (String): The parameter name for error reporting"
    If strValue = "" Or IsNull(strValue) Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.StringNotEmpty", paramName & " cannot be empty or null!"
    End If
End Sub

' Ensures that a string is not whitespace or null
' Args:
'   strValue (String): The string value to check
'   paramName (String): The parameter name for error reporting
Public Sub StringNotWhiteSpace(ByVal strValue As String, ByVal paramName As String)
Attribute StringNotWhiteSpace.VB_Description = "Ensures that a string is not whitespace or null\r\nArgs:\r\n  strValue (String): The string value to check\r\n  paramName (String): The parameter name for error reporting"
    If Trim(strValue) = "" Or IsNull(strValue) Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.StringNotWhiteSpace", paramName & " cannot be whitespace or null!"
    End If
End Sub

' Ensures that an array or collection is not empty
' Args:
'   arr (Variant): The array or collection to check
'   paramName (String): The parameter name for error reporting
Public Sub EnumerableNotEmpty(ByVal arr As Variant, ByVal paramName As String)
Attribute EnumerableNotEmpty.VB_Description = "Ensures that an array or collection is not empty\r\nArgs:\r\n  arr (Variant): The array or collection to check\r\n  paramName (String): The parameter name for error reporting"
    If IsArray(arr) Then
        If UBound(arr) - LBound(arr) + 1 <= 0 Then
            Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.EnumerableNotEmpty", paramName & " cannot be empty!"
        End If
    ElseIf IsObject(arr) Then
        If arr Is Nothing Then
            Err.Raise CommonHResults.ObjectVariableNotSet, "Ensure.EnumerableNotEmpty", paramName & " is Nothing!"
        Else
        End If
    Else
        Err.Raise CommonHResults.InvalidProcedureCall, "Ensure.EnumerableNotEmpty", paramName & " is not an array or collection!"
    End If
End Sub
