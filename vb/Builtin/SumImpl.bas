Attribute VB_Name = "SumImpl"
Option Explicit

' Core implementation for PySum — returns the sum of an iterable of numbers,
' matching Python's sum().
'
' Python semantics:
'   sum([1, 2, 3])        → 6
'   sum([1, 2, 3], 10)    → 16
'   sum([])               → 0
'   sum([], 10)           → 10
'   sum("abc")            → TypeError (can't sum strings)
Public Function PySumCore(ByRef value As Variant, Optional ByVal start As Variant) As Variant

    ' Initialize accumulator with start value (default 0)
    Dim hasStart As Boolean: hasStart = Not IsMissing(start)
    Dim total As Variant
    If hasStart Then
        Builtin.LetSet total, start
    Else
        total = 0
    End If

    Dim iter As IIterator: Set iter = Builtin.PyIter(value)
    Dim item As Variant: Do While Builtin.PyTryNext(iter, item)
        ' Guard: item must be numeric (can't sum strings)
        Ensure.IsTrue IsNumericValue(item), CommonHResults.TypeMismatch, _
            "Builtin.PySum", "unsupported operand type(s) for +: '" & TypeNameOf(item) & "' and 'int'"

        total = total + item
    Loop

    Builtin.LetSet PySumCore, total
End Function

' Returns True if the Variant holds a numeric value.
Private Function IsNumericValue(ByVal v As Variant) As Boolean
    If IsObject(v) Then
        IsNumericValue = False
        Exit Function
    End If

    Select Case VarType(v)
        Case vbEmpty, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsNumericValue = True
        Case Else
            IsNumericValue = False
    End Select
End Function

' Returns a readable type name for error messages (VBA-style type names).
Private Function TypeNameOf(ByVal v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then
            TypeNameOf = "Nothing"
        Else
            TypeNameOf = TypeName(v)
        End If
        Exit Function
    End If

    Select Case VarType(v)
        Case vbEmpty:   TypeNameOf = "Empty"
        Case vbNull:    TypeNameOf = "Null"
        Case vbString:  TypeNameOf = "String"
        Case Else:      TypeNameOf = TypeName(v)
    End Select
End Function