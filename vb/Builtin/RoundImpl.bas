Attribute VB_Name = "RoundImpl"
Option Explicit

' Core implementation for PyRound — rounds a number to ndigits decimal places,
' matching Python's round() using banker's rounding (round half to even).
'
' Python semantics:
'   round(2.675, 2)    → 2.67  (banker's rounding)
'   round(2.5)         → 2     (round half to even)
'   round(3.5)         → 4     (round half to even)
'   round(3.14159, 2)  → 3.14
'   round(3.14159)     → 3
'
' VB6: ndigits is optional. VB6's Round() already uses banker's rounding
' and preserves the input type (Decimal → Decimal, Currency → Currency,
' Double → Double). No custom banker's rounding needed.
Public Function PyRoundCore(ByVal number As Variant, Optional ByVal ndigits As Variant) As Variant
    Dim hasNdigits As Boolean: hasNdigits = Not IsMissing(ndigits)

    If hasNdigits Then
        Ensure.IsTrue IsIntegerValue(ndigits), CommonHResults.TypeMismatch, _
            "Builtin.PyRound", "ndigits must be an integer"

        Dim nd As Long: nd = CLng(ndigits)

        ' VB6 Round(number, ndigits) uses banker's rounding, matching Python.
        ' Type is preserved: Decimal → Decimal, Currency → Currency, Double → Double.
        PyRoundCore = Round(number, nd)
    Else
        ' No ndigits: round to integer
        PyRoundCore = Round(number)
    End If
End Function
