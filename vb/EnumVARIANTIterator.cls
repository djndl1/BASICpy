VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumVARIANTIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterable
Implements IIterator

' --- IEnumVARIANT::Next HRESULT values ---
Private Const S_OK As Long = 0
Private Const S_FALSE As Long = 1

Private Const VT_BYREF As Integer = &H4000

' --- VTable offsets for IEnumVARIANT (32-bit, 4 bytes per vtable entry) ---
' IUnknown:  QueryInterface(0), AddRef(4), Release(8)
' IEnumVARIANT: Next(12), Skip(16), Reset(20), Clone(24)
Private Const VF_NEXT As Long = 0

Private Const CC_STDCALL As Long = 4

Private Declare Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As Long, _
    ByRef pvargResult As Variant) As Long

Private m_enumObj As IEnumVARIANT     ' IEnumVARIANT (held as IUnknown via _NewEnum)
Private m_exhausted As Boolean  ' True once S_FALSE has been returned

' Initialises the iterator by getting the IEnumVARIANT from the source object.
' source must support _NewEnum (the COM enumerator behind For Each).
Public Sub Init(ByRef source As Object)
    If TypeOf source Is Collection Then
        Dim col As Collection: Set col = source
            Set m_enumObj = col.[_NewEnum]
        Exit Sub
    End If
    On Error Resume Next
    Dim enumResult As Variant
    enumResult = CallByDispId(source, -4, 2)  ' 2 = VbGet = DISPATCH_PROPERTYGET
    If VarType(enumResult) = vbObject Then
        Set m_enumObj = enumResult
    End If
    If Err.Number <> 0 Or m_enumObj Is Nothing Then
        Err.Raise 5, "EnumVARIANTIterator.Init", _
            "Object does not support IEnumVARIANT (_NewEnum)"
    End If
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set m_enumObj = Nothing
End Sub

Private Function IIterator_NextItem() As Variant
    If m_exhausted Then
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
        Exit Function
    End If

    Dim vCelt As Variant: vCelt = CLng(1)
    ReDim rgVar(0 To vCelt - 1) As Variant
    Dim pvValue As Variant: pvValue = VarPtr(rgVar(0))
    Dim vFetched As Long: vFetched = CLng(0)
    Dim pvFetched As Variant: pvFetched = VarPtr(vFetched)
    ' Parameter value pointers
    Dim pv(0 To 2) As Long
    Dim vt(0 To 2) As Integer
    pv(0) = VarPtr(vCelt)
    vt(0) = vbLong
    pv(1) = VarPtr(pvValue)
    vt(1) = vbVariant Or VT_BYREF
    pv(2) = VarPtr(pvFetched)
    vt(2) = vbLong Or VT_BYREF

    ' Dispatch through vtable: IEnumVARIANT::Next at offset 12
    Dim result As Variant
    Dim hr As Long
    hr = DispCallFunc(ObjPtr(m_enumObj), 12, CC_STDCALL, vbLong, 3, _
                      vt(0), pv(0), result)

    If hr = 0 And VarType(result) = vbLong Then
        If CLng(result) = S_OK And vFetched >= 1 Then
            IIterator_NextItem = vValue
        Else
            m_exhausted = True
            IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
        End If
    Else
        m_exhausted = True
        IIterator_NextItem = CVErr(StopIterationCode.StopIteration)
    End If
End Function

Private Function IIterator_Iter() As IIterator
    Set IIterator_Iter = Me
End Function

Private Function IIterable_Iter() As IIterator
    Set IIterable_Iter = Me
End Function
