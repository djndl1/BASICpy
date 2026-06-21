Attribute VB_Name = "PyBoolImpl"
Option Explicit

' Returns the truth value of a value, matching Python's bool() truth testing.
' https://docs.python.org/3/library/stdtypes.html#truth-value-testing
' Dispatch: primitives > PyLen > IBoolean (__bool__) > ILength (__len__) > True
Public Function PyBoolCore(Optional ByRef value As Variant) As Boolean
    ' 0. Nothing → False (Python: bool(None) → False)
    If IsObject(value) Then
        If value Is Nothing Then
            PyBool = False
            Exit Function
        End If
    End If
    
    ' 1. No argument (missing), Empty, or Null → False
    ' IsMissing throws on undimensioned arrays (NULL SAFEARRAY), so we guard it.
    If Not IsArray(value) Then
        If IsMissing(value) Then
            PyBool = False
            Exit Function
        End If
    End If
    
    ' Empty or Null → False (safe on any Variant type)
    If IsEmpty(value) Or IsNull(value) Then
        PyBool = False
        Exit Function
    End If
    
    ' 2. Boolean → value itself
    If VarType(value) = vbBoolean Then
        PyBool = value
        Exit Function
    End If
    
    ' 3. Numeric types → non-zero is truthy
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            PyBool = (value <> 0)
            Exit Function
        Case vbDate
            PyBool = (CDbl(value) <> 0)
            Exit Function
    End Select
    
    ' 4. Use PyLenImpl for anything length-detectable
    Dim length As Long
    length = PyLenImpl(value)
    If length <> -1 Then
        PyBool = (length > 0)
        Exit Function
    End If
    
    ' If PyLenImpl failed but value IS an array → undimensioned/empty → False
    If IsArray(value) Then
        PyBool = False
        Exit Function
    End If
    
    ' 5. Object — protocol dispatch
    If IsObject(value) Then
        If value Is Nothing Then
            PyBool = False
            Exit Function
        End If
        
        ' __bool__ protocol (IBoolean)
        If TypeOf value Is IBoolean Then
            Dim boolObj As IBoolean
            Set boolObj = value
            PyBool = boolObj.Bool()
            Exit Function
        End If
        
        ' __len__ protocol (ILength)
        If TypeOf value Is ILength Then
            Dim lenObj As ILength
            Set lenObj = value
            PyBool = (lenObj.Length() > 0)
            Exit Function
        End If
        
        ' Default: objects are truthy
        PyBool = True
        Exit Function
    End If
    
    ' 7. Fallback — try CBool, default True
    On Error Resume Next
    PyBool = CBool(value)
    If Err.Number <> 0 Then PyBool = True
    On Error GoTo 0
End Function
