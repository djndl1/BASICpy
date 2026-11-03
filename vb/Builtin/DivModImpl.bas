Attribute VB_Name = "DivModImpl"
Option Explicit

' Core implementation for PyDivMod — returns the quotient and remainder,
' matching Python's divmod().
'
' Python semantics:
'   divmod(7, 3)   → (2, 1)        both int → int result
'   divmod(7, -3)  → (-3, -2)
'   divmod(-7, 3)  → (-3, 2)
'   divmod(-7, -3) → (2, -1)
'   divmod(7.5, 2) → (3.0, 1.5)    any float → float result
'   divmod(7.0, 3) → (2.0, 1.0)
'
' Returns a 2-element Variant array: (quotient, remainder).
'
' Type preservation:
'   - Decimal vs anything numeric → Decimal arithmetic (28-digit precision)
'   - Currency vs anything numeric → Currency arithmetic
'   - Both strict integer (Byte/Integer/Long/Boolean) → Integer arithmetic
'   - Any float (Single/Double) or mixed int/float → Double
Public Function PyDivModCore(ByVal a As Variant, ByVal b As Variant) As Variant
    Ensure.IsTrue b <> 0, CommonHResults.DivisionByZero, _
        "Builtin.PyDivMod", "integer division or modulo by zero"

    Dim result(0 To 1) As Variant
    Dim aType As VbVarType
    Dim bType As VbVarType
    aType = VarType(a)
    bType = VarType(b)

    ' Python's divmod uses floor division: q = floor(a / b), r = a - q * b

    ' Decimal path: any operand is Decimal → Decimal arithmetic (coerce other)
    If aType = vbDecimal Or bType = vbDecimal Then
        Dim aDec As Variant: aDec = a
        Dim bDec As Variant: bDec = b
        If aType <> vbDecimal Then aDec = CDec(a)
        If bType <> vbDecimal Then bDec = CDec(b)
        Dim qDec As Variant: qDec = Int(aDec / bDec)
        result(0) = qDec
        result(1) = aDec - qDec * bDec
        Builtin.LetSet PyDivModCore, result
        Exit Function
    End If

    ' Currency path: any operand is Currency → Currency arithmetic (coerce other)
    If aType = vbCurrency Or bType = vbCurrency Then
        Dim aCur As Variant: aCur = a
        Dim bCur As Variant: bCur = b
        If aType <> vbCurrency Then aCur = CCur(a)
        If bType <> vbCurrency Then bCur = CCur(b)
        Dim qCur As Variant: qCur = Int(aCur / bCur)
        result(0) = qCur
        result(1) = aCur - qCur * bCur
        Builtin.LetSet PyDivModCore, result
        Exit Function
    End If

    ' Strict integer path: both operands are Byte/Integer/Long/Boolean → Integer result
    If IsStrictIntegerType(aType) And IsStrictIntegerType(bType) Then
        Dim q As Long: q = CLng(Int(a / b))   ' floor division as Long
        result(0) = q
        result(1) = a - q * b
        Builtin.LetSet PyDivModCore, result
        Exit Function
    End If

    ' Double fallback: any float (Single/Double) or mixed int/float → Double result
    Dim qDbl As Double: qDbl = Int(CDbl(a) / CDbl(b))
    Dim rDbl As Double: rDbl = CDbl(a) - qDbl * CDbl(b)
    result(0) = qDbl
    result(1) = rDbl

    Builtin.LetSet PyDivModCore, result
End Function
