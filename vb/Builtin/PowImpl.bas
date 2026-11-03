Attribute VB_Name = "PowImpl"
Option Explicit

' Core implementation for PyPow — returns base raised to the power exp,
' matching Python's pow().
'
' Python semantics:
'   pow(2, 3)        → 8
'   pow(2, 3, 5)     → 3   (modular exponentiation)
'   pow(2, -1)       → 0.5
'   pow(2, 0)        → 1
'   pow(38, -1, 97)  → 23  (modular inverse, Python 3.8+)
'
' VB6: exp and mod are optional. mod requires all args to be integers.
'      Negative exp with mod computes modular inverse (requires coprime).
Public Function PyPowCore(ByVal base As Variant, ByVal exp As Variant, Optional ByVal modu As Variant) As Variant
    Dim hasMod As Boolean: hasMod = Not IsMissing(modu)

    ' Validate mod requirements (Python: 3rd arg only allowed for integers)
    If hasMod Then
        Ensure.IsTrue IsIntegerValue(modu), CommonHResults.InvalidProcedureCall, _
            "Builtin.PyPow", "pow() 3rd argument not allowed unless all arguments are integers"
        Ensure.IsTrue modu <> 0, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyPow", "pow() 3rd argument cannot be 0"
        Ensure.IsTrue IsIntegerValue(base), CommonHResults.InvalidProcedureCall, _
            "Builtin.PyPow", "pow() 3rd argument not allowed unless all arguments are integers"
        Ensure.IsTrue IsIntegerValue(exp), CommonHResults.InvalidProcedureCall, _
            "Builtin.PyPow", "pow() 3rd argument not allowed unless all arguments are integers"

        ' Negative exponent with mod: modular inverse (Python 3.8+)
        ' pow(base, -exp, mod) = pow(mod_inverse(base, mod), exp, mod)
        ' Requires base and mod to be coprime.
        If exp < 0 Then
            Dim invBase As Variant
            invBase = ModInverse(base, modu)
            Dim posExp As Variant: posExp = -exp
            PyPowCore = ModPow(invBase, posExp, modu)
            Exit Function
        End If

        ' Positive exponent with mod: efficient modular exponentiation
        ' Uses square-and-multiply with intermediate mod to avoid overflow.
        PyPowCore = ModPow(base, exp, modu)
        Exit Function
    End If

    ' No mod: simple power (VB6 ^ operator handles int, float, negative exp)
    On Error Resume Next
    PyPowCore = base ^ exp
    If Err.number <> 0 Then
        Err.Raise CommonHResults.InvalidProcedureCall, "Builtin.PyPow", _
            "pow() computation failed"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
End Function

' Efficient modular exponentiation: (base^exp) mod modu.
' Uses square-and-multiply with intermediate mod reduction to avoid overflow.
' All operands are integers; result is integer.
Private Function ModPow(ByVal base As Variant, ByVal exp As Variant, ByVal modu As Variant) As Variant
    Dim result As Variant: result = 1
    Dim b As Variant: b = base Mod modu
    ' VB6 Mod can return negative remainder; normalize to [0, modu) like Python %
    If b < 0 Then b = b + modu
    Dim e As Variant: e = exp
    Dim one As Variant: one = 1
    Dim two As Variant: two = 2

    Do While e > 0
        If e Mod two = one Then
            result = (result * b) Mod modu
        End If
        e = e \ two
        b = (b * b) Mod modu
    Loop

    ModPow = result
End Function

' Modular multiplicative inverse: returns x such that (base * x) mod modu = 1.
' Uses extended Euclidean algorithm. Requires base and modu to be coprime.
' Raises error if gcd(base, modu) <> 1 (no inverse exists).
Private Function ModInverse(ByVal base As Variant, ByVal modu As Variant) As Variant
    Dim oldR As Variant: oldR = base
    Dim r As Variant: r = modu
    Dim oldS As Variant: oldS = 1
    Dim s As Variant: s = 0
    Dim zero As Variant: zero = 0

    Do While r <> zero
        Dim quotient As Variant: quotient = oldR \ r
        Dim temp As Variant
        temp = oldR - quotient * r: oldR = r: r = temp
        temp = oldS - quotient * s: oldS = s: s = temp
    Loop

    ' oldR is gcd; must be 1 for inverse to exist
    Ensure.IsTrue oldR = 1, CommonHResults.InvalidProcedureCall, _
        "Builtin.PyPow", "base is not relatively prime to the modulus"

    ' oldS is the inverse, may be negative; normalize to [0, modu)
    Dim result As Variant: result = oldS Mod modu
    If result < 0 Then result = result + modu
    ModInverse = result
End Function
