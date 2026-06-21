Attribute VB_Name = "TupleImpl"
Option Explicit

' Returns a new tuple (immutable sequence) wrapping the given value.
'   PyTuple()             → ()
'   PyTuple(arr)          → (arr(0), arr(1), ...) — copy of array
'   PyTuple(tupleObj)     → tupleObj (returned unchanged)
'   PyTuple("str")        → ('s', 't', 'r') — iterate chars
'   PyTuple(iterable)     → collect all items, wrap as tuple
Public Function PyTupleCore(Optional ByRef value As Variant) As Variant
    Dim tup As Tuple
    Dim objVal As Object
    
    ' No argument → empty tuple
    If IsMissing(value) Then
        Set tup = New Tuple
        tup.Init Array()
        Set PyTuple = tup
        Exit Function
    End If
    
    ' Already a Tuple → return as-is (same object)
    If IsObject(value) Then
        Set objVal = value
        If TypeOf objVal Is Tuple Then
            Set PyTuple = objVal
            Exit Function
        End If
    End If
    
    ' Array → copy into new Tuple
    If IsArray(value) Then
        Set tup = New Tuple
        tup.Init value
        Set PyTuple = tup
        Exit Function
    End If
    
    ' String → iterate characters
    If VarType(value) = vbString Then
        Dim str As String
        str = value
        Dim strLen As Long: strLen = Len(str)
        Dim strArr() As Variant
        Dim idx As Long
        If strLen > 0 Then
            ReDim strArr(0 To strLen - 1)
            For idx = 0 To strLen - 1
                strArr(idx) = Mid$(str, idx + 1, 1)
            Next idx
        Else
            strArr = Array()
        End If
        Set tup = New Tuple
        tup.Init strArr
        Set PyTuple = tup
        Exit Function
    End If
    
    ' Iterable (IIterable / IIterator / COM enumeration) → collect all items
    Dim iter As IIterator
    Set iter = PyIter(value)
    
    ' Collect into a growable array using Collection as a dynamic buffer
    Dim buffer As New Collection
    Dim item As Variant
    Do While PyTryNext(iter, item)
        buffer.Add item
    Loop
    
    ' Convert collection to array
    Dim bufLen As Long: bufLen = buffer.count
    Dim resultArr() As Variant
    Dim pos As Long
    If bufLen > 0 Then
        ReDim resultArr(0 To bufLen - 1)
        For pos = 1 To bufLen
            resultArr(pos - 1) = buffer(pos)
        Next pos
    Else
        resultArr = Array()
    End If
    
    Set tup = New Tuple
    tup.Init resultArr
    Set PyTuple = tup
End Function
