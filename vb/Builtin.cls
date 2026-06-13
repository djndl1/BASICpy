VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Builtin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function PyAll(ByRef arr As Variant, Optional ByVal predicate As Variant) As Boolean
    If IsMissing(predicate) Then
        Set predicate = IsTruthy
    End If

    If IsArray(arr) Then
        PyAll = IterateArray(arr, predicate, True, True)
    ElseIf TypeName(arr) = "Collection" Then
        PyAll = IterateCollection(arr, predicate, True, True)
    ElseIf IsObject(arr) And IsIEnumVARAINT(arr) Then
        PyAll = IterateIEnumVariant(arr, predicate, True, True)
    Else
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyAll", "argument is neither an array nor a Collection nor an enumerable object"
    End If
End Function

Public Function PyAny(ByRef arr As Variant, Optional ByVal predicate As Variant) As Boolean
    If IsMissing(predicate) Then
        Set predicate = IsTruthy
    End If

    If IsArray(arr) Then
        PyAny = IterateArray(arr, predicate, False, False)
    ElseIf TypeName(arr) = "Collection" Then
        PyAny = IterateCollection(arr, predicate, False, False)
    ElseIf IsObject(arr) And IsIEnumVARAINT(arr) Then
        PyAny = IterateIEnumVariant(arr, predicate, False, False)
    Else
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyAny", "argument is neither an array nor a Collection nor an enumerable object"
    End If
End Function

Private Function IterateArray(ByRef arr As Variant, _
                             ByVal predicate As IPredicate, _
                             ByVal modeIsAll As Boolean, _
                             ByVal emptyDefault As Boolean) As Boolean
    Dim i As Long
    If LBound(arr) > UBound(arr) Then
        IterateArray = emptyDefault
        Exit Function
    End If
    For i = LBound(arr) To UBound(arr)
        If ReduceStep(predicate.Test(arr(i)), modeIsAll, emptyDefault, IterateArray) Then Exit Function
    Next i
    IterateArray = emptyDefault
End Function

Private Function IterateCollection(ByVal col As Collection, _
                                  ByVal predicate As IPredicate, _
                                  ByVal modeIsAll As Boolean, _
                                  ByVal emptyDefault As Boolean) As Boolean
    Dim i As Long
    If col.Count = 0 Then
        IterateCollection = emptyDefault
        Exit Function
    End If
    For i = 1 To col.Count
        If ReduceStep(predicate.Test(col(i)), modeIsAll, emptyDefault, IterateCollection) Then Exit Function
    Next i
    IterateCollection = emptyDefault
End Function

Private Function IterateIEnumVariant(ByVal obj As Object, _
                                     ByVal predicate As IPredicate, _
                                     ByVal modeIsAll As Boolean, _
                                     ByVal emptyDefault As Boolean) As Boolean
    Dim element As Variant
    For Each element In obj
        If ReduceStep(predicate.Test(element), modeIsAll, emptyDefault, IterateIEnumVariant) Then Exit Function
    Next element
    IterateIEnumVariant = emptyDefault
End Function

Private Function ReduceStep(ByVal passes As Boolean, _
                            ByVal modeIsAll As Boolean, _
                            ByVal emptyDefault As Boolean, _
                            ByRef outResult As Boolean) As Boolean
    If modeIsAll Then
        If Not passes Then
            outResult = False
            ReduceStep = True
        End If
    Else
        If passes Then
            outResult = True
            ReduceStep = True
        End If
    End If
End Function

Private Function IsIEnumVARAINT(ByVal obj As Object) As Boolean
    On Error Resume Next
    Dim enumObj As Object
    Set enumObj = CallByName(obj, "_NewEnum", VbGet)
    If Err.Number <> 0 Then
        IsIEnumVARAINT = False
    Else
        IsIEnumVARAINT = Not (enumObj Is Nothing)
    End If
    On Error GoTo 0
End Function
