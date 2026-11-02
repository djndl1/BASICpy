Attribute VB_Name = "ChrOrdImpl"
Option Explicit

' Core implementation for PyChr — converts an integer to a character,
' matching Python's chr().
'
' Python semantics:
'   chr(65)        → 'A'
'   chr(97)        → 'a'
'   chr(8364)      → '€'   (U+20AC, BMP)
'   chr(128512)    → '😀'  (U+1F600, supplementary plane via surrogate pair)
'   chr(0)         → ChrW$(0)
'   Range: 0 to 0x10FFFF (Python 3 full Unicode range)
'
' VB6: Uses ChrW$() for Unicode (UTF-16). Supplementary characters
'      (0x10000+) are encoded as surrogate pairs (2 UTF-16 code units).
' Conversion is delegated to PyInt which handles IIndex objects,
' numeric types, and float truncation via Fix().
Public Function PyChrCore(ByVal value As Variant) As String
    Dim code As Long

    ' Convert via PyInt: handles IIndex, numeric types, float truncation.
    ' Catch any type errors and re-raise with chr-specific context.
    On Error Resume Next
    code = Builtin.PyInt(value)
    If Err.number <> 0 Then
        Dim errNum As Long: errNum = Err.number
        Dim errDesc As String: errDesc = Err.Description
        On Error GoTo 0
        Err.Raise errNum, "Builtin.PyChr", errDesc
        Exit Function
    End If
    On Error GoTo 0

    ' Python allows 0 to 0x10FFFF
    Ensure.IsTrue code >= 0 And code <= &H10FFFF, CommonHResults.InvalidProcedureCall, _
        "Builtin.PyChr", "chr() argument must be in range(0x110000)"

    PyChrCore = ChrCodeToString(code)
End Function

' Converts a Unicode code point (0-0x10FFFF) to a UTF-16 VB6 String.
' BMP code points (0-65535) become a single ChrW$ character.
' Supplementary code points (0x10000-0x10FFFF) become a surrogate pair.
Private Function ChrCodeToString(ByVal code As Long) As String
    If code <= &HFFFF Then
        ChrCodeToString = ChrW$(code)
    Else
        Dim offset As Long: offset = code - &H10000
        Dim highSurrogate As Long: highSurrogate = &HD800 + (offset \ &H400)
        Dim lowSurrogate As Long: lowSurrogate = &HDC00 + (offset Mod &H400)
        ChrCodeToString = ChrW$(highSurrogate) & ChrW$(lowSurrogate)
    End If
End Function

' Core implementation for PyOrd — returns the integer code of a character,
' matching Python's ord().
'
' Python semantics:
'   ord('A')     → 65
'   ord('a')     → 97
'   ord('€')     → 8364   (U+20AC, BMP)
'   ord('😀')    → 128512 (U+1F600, supplementary via surrogate pair)
'   ord('')      → raises TypeError (empty string)
'   ord('ab')    → raises TypeError (not a single character)
'
' VB6: Uses AscW() to read UTF-16 code units. A supplementary character
'      stored as a surrogate pair (2 code units) is decoded back to the
'      full Unicode code point.
Public Function PyOrdCore(ByVal value As String) As Long
    Dim slen As Long: slen = Len(value)

    Ensure.IsTrue slen > 0, CommonHResults.InvalidProcedureCall, _
        "Builtin.PyOrd", "ord() expected a character, but string is empty"
    Dim code As Long
    ' Single UTF-16 code unit (BMP character)
    If slen = 1 Then
        code = CLng(AscW(value))
        ' AscW returns signed Integer; convert &H8000-&HFFFF to unsigned
        If code < 0 Then code = code + 65536
        PyOrdCore = code
        Exit Function
    End If

    ' Two UTF-16 code units — could be a surrogate pair (1 supplementary char)
    If slen = 2 Then
        Dim firstCode As Long: firstCode = CLng(AscW(Left$(value, 1)))
        If firstCode < 0 Then firstCode = firstCode + 65536
        Dim secondCode As Long: secondCode = CLng(AscW(Mid$(value, 2, 1)))
        If secondCode < 0 Then secondCode = secondCode + 65536

        ' Check for valid surrogate pair: high (0xD800-0xDBFF) + low (0xDC00-0xDFFF)
        If firstCode >= &HD800 And firstCode <= &HDBFF And _
           secondCode >= &HDC00 And secondCode <= &HDFFF Then
            code = &H10000 + ((firstCode - &HD800) * &H400) + (secondCode - &HDC00)
            PyOrdCore = code
            Exit Function
        End If
    End If

    ' Any other length or invalid surrogate pair → error
    Err.Raise CommonHResults.InvalidProcedureCall, "Builtin.PyOrd", _
        "ord() expected a single character, but string of length " & slen & " found"
End Function
