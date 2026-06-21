Attribute VB_Name = "VBArrayCoreImpl"
Option Explicit

' Core implementation for PyVBArray — returns a new VBArray from the given value.
Public Function PyVBArrayCore(Optional ByRef value As Variant) As VBArray
    Dim vba As VBArray
    Dim objVal As Object
    
    ' No argument → empty
    If IsMissing(value) Then
        Set PyVBArrayCore = New VBArray
        Exit Function
    End If
    
    ' Already a VBArray → use Copy (single array copy vs O(n) Append)
    If IsObject(value) Then
        Set objVal = value
        If TypeOf objVal Is VBArray Then
            Dim src As VBArray
            Set src = objVal
            Set PyVBArrayCore = src.Copy()
            Exit Function
        End If
    End If
    
    ' Array → copy into new VBArray
    If IsArray(value) Then
        Set vba = New VBArray
        vba.Init value
        Set PyVBArrayCore = vba
        Exit Function
    End If
    
    ' String → iterate characters
    If VarType(value) = vbString Then
        Set vba = New VBArray
        Dim str As String: str = value
        Dim strLen As Long: strLen = Len(str)
        Dim idx As Long
        For idx = 1 To strLen
            vba.Append Mid$(str, idx, 1)
        Next idx
        Set PyVBArrayCore = vba
        Exit Function
    End If
    
    ' Iterable → collect all items
    Dim iter As IIterator: Set iter = PyIter(value)
    Set vba = New VBArray
    Dim item As Variant: Do While PyTryNext(iter, item)
        vba.Append item
    Loop
    Set PyVBArrayCore = vba
End Function
