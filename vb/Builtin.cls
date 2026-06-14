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
        PyAll = PyAllArray(arr, predicate)
    ElseIf TypeName(arr) = "Collection" Then
        PyAll = PyAllCollection(arr, predicate)
    ElseIf IsObject(arr) And IsIEnumVARAINT(arr) Then
        PyAll = PyAllIEnumVariant(arr, predicate)
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
        PyAny = PyAnyArray(arr, predicate)
    ElseIf TypeName(arr) = "Collection" Then
        PyAny = PyAnyCollection(arr, predicate)
    ElseIf IsObject(arr) And IsIEnumVARAINT(arr) Then
        PyAny = PyAnyIEnumVariant(arr, predicate)
    Else
        Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
            "Builtin.PyAny", "argument is neither an array nor a Collection nor an enumerable object"
    End If
End Function

Public Function PyAbs(ByVal value As Variant) As Variant
    If IsObject(value) Then
        If TypeOf value Is IAbsolute Then
            Dim absImpl As IAbsolute
            Set absImpl = value
            PyAbs = absImpl.Absolute()
        Else
            Ensure.IsTrue False, CommonHResults.InvalidProcedureCall, _
                "Builtin.PyAbs", "object does not implement IAbsolute"
        End If
    Else
        PyAbs = Abs(value)
    End If
End Function

Private Function PyAllArray(ByRef arr As Variant, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Not predicate.Test(arr(i)) Then
            PyAllArray = False
            Exit Function
        End If
    Next i
    PyAllArray = True
End Function

Private Function PyAllCollection(ByVal col As Collection, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If Not predicate.Test(col(i)) Then
            PyAllCollection = False
            Exit Function
        End If
    Next i
    PyAllCollection = True
End Function

Private Function PyAllIEnumVariant(ByVal obj As Object, ByVal predicate As IPredicate) As Boolean
    Dim element As Variant
    For Each element In obj
        If Not predicate.Test(element) Then
            PyAllIEnumVariant = False
            Exit Function
        End If
    Next element
    PyAllIEnumVariant = True
End Function

Private Function PyAnyArray(ByRef arr As Variant, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If predicate.Test(arr(i)) Then
            PyAnyArray = True
            Exit Function
        End If
    Next i
    PyAnyArray = False
End Function

Private Function PyAnyCollection(ByVal col As Collection, ByVal predicate As IPredicate) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If predicate.Test(col(i)) Then
            PyAnyCollection = True
            Exit Function
        End If
    Next i
    PyAnyCollection = False
End Function

Private Function PyAnyIEnumVariant(ByVal obj As Object, ByVal predicate As IPredicate) As Boolean
    Dim element As Variant
    For Each element In obj
        If predicate.Test(element) Then
            PyAnyIEnumVariant = True
            Exit Function
        End If
    Next element
    PyAnyIEnumVariant = False
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
