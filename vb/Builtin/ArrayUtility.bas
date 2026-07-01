Attribute VB_Name = "ArrayUtility"
Option Explicit

' Normalizes a 0-based index (supports negative) and returns the storage index (lb + Index).
' Raises SubscriptOutOfRange if index is out of bounds.
' allowAppend=True permits Index=length (for insert-at-end semantics).
Public Function ArrayResolveIndex(ByVal Index As Long, ByVal length As Long, _
                                  ByVal lb As Long, _
                                  Optional ByVal allowAppend As Boolean = False) As Long
    If Index < 0 Then Index = length + Index

    Dim max As Long
    If allowAppend Then max = length Else max = length - 1

    If Index < 0 Or Index > max Then
        Ensure.IsTrue False, CommonHResults.SubscriptOutOfRange, _
            "ArrayResolveIndex", "index (" & Index & ") out of range"
    End If

    ArrayResolveIndex = lb + Index
End Function

' Returns the number of occurrences of value in arr(lb To lb + length - 1).
Public Function ArrayCountOf(ByRef arr As Variant, ByVal lb As Long, _
                             ByVal length As Long, ByRef value As Variant) As Long
    Dim i As Long
    For i = lb To lb + length - 1
        If arr(i) = value Then ArrayCountOf = ArrayCountOf + 1
    Next i
End Function

' Returns the index (0-based) of the first occurrence of value within [start, stopAt).
' Returns -1 if value is not found.
' start defaults to 0, stopAt defaults to -1 (meaning all elements).
' Supports negative start/stopAt (relative to length).
Public Function ArrayIndexOf(ByRef arr As Variant, ByVal lb As Long, _
                             ByVal length As Long, ByRef value As Variant, _
                             Optional ByVal start As Long = 0, _
                             Optional ByVal stopAt As Long = -1) As Long

    ' Default stopAt to length (entire range)
    If stopAt = -1 Then stopAt = length

    ' Normalize negative indices
    If start < 0 Then start = length + start
    If start < 0 Then start = 0
    If stopAt < 0 Then stopAt = length + stopAt
    If stopAt > length Then stopAt = length

    Dim i As Long
    For i = lb + start To lb + stopAt - 1
        If arr(i) = value Then
            ArrayIndexOf = i - lb
            Exit Function
        End If
    Next i

    ArrayIndexOf = -1
End Function

' Reverses elements in arr(lb To lb + length - 1) in-place.
Public Sub ReverseInPlace(ByRef arr As Variant, ByVal lb As Long, ByVal length As Long)
    Dim i As Long: i = lb
    Dim j As Long: j = lb + length - 1
    Dim t As Variant
    While i < j
        t = arr(i)
        arr(i) = arr(j)
        arr(j) = t
        i = i + 1
        j = j - 1
    Wend
End Sub