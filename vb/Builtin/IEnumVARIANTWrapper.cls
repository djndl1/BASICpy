VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEnumVARIANTWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Wraps the IEnumVARIANT COM interface behind For Each / _NewEnum,
' exposing Next, Skip, Reset, and Clone through VB6-friendly methods.
' Uses DispCall.DispCallFunc for raw vtable dispatch (slots 3–6 × DispCall.POINTER_LENGTH).

Private m_enum As IEnumVARIANT

' --- Lifecycle ---

' Initialises the wrapper from an existing IEnumVARIANT (IUnknown pointer).
Public Sub InitFromEnumerator(ByRef enumObj As IUnknown)
    Set m_enum = enumObj
End Sub

' Initialises the wrapper from any _NewEnum-supporting object.
' Calls TryGetEnum to obtain the IEnumVARIANT internally.
Public Sub InitFromEnumerable(ByRef source As Object)
    Dim enumObj As IUnknown
    Dim callResult As Long: callResult = DispCall.TryGetEnum(source, enumObj)

    If callResult <> 0 Or enumObj Is Nothing Then
        Err.Raise 5, "IEnumVARIANTWrapper.InitFromEnumerable", _
            "Object does not support IEnumVARIANT (_NewEnum): internal error " & callResult
    End If
    Set m_enum = enumObj
End Sub

Private Sub Class_Terminate()
    Set m_enum = Nothing
End Sub

' --- Public API ---

' Fetches up to count items from the enumerator.
' Returns a Variant() array — empty array means exhaustion.
' Raises a VB6 error on COM failure (Option A: hide HRESULTs from callers).
Public Function NextItems(ByVal count As Long) As Variant()
    ' Guard against degenerate call
    If count <= 0 Then
        NextItems = EmptyVariantArray
        Exit Function
    End If

    ' Pre-allocate VARIANT buffer — COM writes contiguous Variant structs
    ReDim buffer(0 To count - 1) As Variant
    Dim result As Variant
    Dim vFetched As Long
    Dim hr As Long: hr = NextRaw(count, VarPtr(buffer(0)), result, vFetched)

    ' COM failure should propagate as a VB6 error
    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.NextItems", _
            "IEnumVARIANT::Next returned error"
        Exit Function
    End If

    ' Guards: result must be S_OK (0) with fetched > 0
    If VarType(result) <> vbLong Then
        NextItems = EmptyVariantArray
        Exit Function
    End If
    If CLng(result) <> CommonHResults.Success Then
        NextItems = EmptyVariantArray
        Exit Function
    End If
    If vFetched <= 0 Then
        NextItems = EmptyVariantArray
        Exit Function
    End If

    ' Here: S_OK with fetched > 0
    If vFetched = count Then
        NextItems = buffer
        Exit Function
    End If

    ' S_FALSE — partial fetch, truncate buffer to actual count
    ReDim Preserve buffer(0 To vFetched - 1)
    NextItems = buffer
End Function

' ─── NextItem (single-element fast path) ───

' Fetches a single item from the enumerator.
' Returns the item as a Variant, or Empty if exhausted.
' Uses a stack-allocated VARIANT instead of a heap array.
Public Function NextItem() As Variant
    Dim buffer0 As Variant    ' stack VARIANT — Next writes result here
    Dim result As Variant
    Dim vFetched As Long
    Dim hr As Long: hr = NextRaw(1, VarPtr(buffer0), result, vFetched)

    ' COM failure should propagate as a VB6 error
    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.NextItem", _
            "IEnumVARIANT::Next returned error"
        Exit Function
    End If

    ' Guards: result must be S_OK (0) with fetched > 0
    If VarType(result) <> vbLong Then Exit Function
    If CLng(result) <> CommonHResults.Success Then Exit Function
    If vFetched <= 0 Then Exit Function

    ' Here: S_OK with fetched > 0
    NextItem = buffer0
End Function

' ─── Private helper: raw IEnumVARIANT::Next vtable call ───

' Sets up the 3-parameter DispCallFunc for IEnumVARIANT::Next and returns
' the HRESULT, the result Variant, and the number of items fetched.
' Callers must check the returned HRESULT and handle their own errors.
Private Function NextRaw(ByVal celt As Long, ByVal rgvarPtr As Long, _
    ByRef outResult As Variant, ByRef outFetched As Long) As Long
    Dim vCelt As Variant: vCelt = CLng(celt)
    Dim pvBuffer As Variant: pvBuffer = rgvarPtr
    Dim vFetched As Long: vFetched = 0
    Dim pvFetched As Variant: pvFetched = VarPtr(vFetched)

    ' IEnumVARIANT::Next(ULONG celt, VARIANT *rgvar, ULONG *pceltFetched)
    Dim vt(0 To 2) As Integer
    Dim pv(0 To 2) As Long
    pv(0) = VarPtr(vCelt)
    vt(0) = vbLong
    pv(1) = VarPtr(pvBuffer)
    vt(1) = vbVariant Or DispCall.VT_BYREF
    pv(2) = VarPtr(pvFetched)
    vt(2) = vbLong Or DispCall.VT_BYREF

    Dim result As Variant
    Dim hr As Long
    hr = DispCall.DispCallFunc(ObjPtr(m_enum), 3 * DispCall.POINTER_LENGTH, _
        DispCall.CALLCONV.CC_STDCALL, vbLong, 3, _
        vt(0), pv(0), result)

    outResult = result
    outFetched = vFetched
    NextRaw = hr
End Function

' Returns a zero-length Variant() array — used as exhaustion sentinel and
' degenerate-case guard. An undimensioned dynamic array is implicitly empty.
Private Function EmptyVariantArray() As Variant()
    'ReDim arr(1 To 0)
    EmptyVariantArray = Array()
End Function

' Skips count items in the enumeration.
' After Skip, the next NextItems call returns items AFTER the skipped range.
Public Sub Skip(ByVal count As Long)
    If count <= 0 Then Exit Sub

    Dim vCelt As Variant: vCelt = CLng(count)
    Dim pv As Long: pv = VarPtr(vCelt)
    Dim vt As Integer: vt = vbLong
    Dim result As Variant
    Dim hr As Long: hr = DispCall.DispCallFunc(ObjPtr(m_enum), 4 * DispCall.POINTER_LENGTH, _
                                               DispCall.CALLCONV.CC_STDCALL, vbLong, 1, _
                                               vt, pv, result)

    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.Skip", _
            "IEnumVARIANT::Skip returned error"
    End If
End Sub

' Resets the enumerator back to the beginning.
' After Reset, the next NextItems call returns items from the start.
Public Sub Reset()
    Dim vt As Integer
    Dim pv As Long
    Dim result As Variant
    Dim hr As Long: hr = DispCall.DispCallFunc(ObjPtr(m_enum), 5 * DispCall.POINTER_LENGTH, _
                                               DispCall.CALLCONV.CC_STDCALL, vbLong, 0, _
                                               vt, pv, result)

    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.Reset", _
            "IEnumVARIANT::Reset returned error"
    End If
End Sub

Public Function Clone() As IEnumVARIANTWrapper
    Dim vtClone As Integer: vtClone = vbDataObject Or DispCall.VT_BYREF
    Dim ppEnum As IUnknown
    Dim pvClone As Long: pvClone = VarPtr(ppEnum)

    Dim result As Variant
    Dim hr As Long: hr = DispCall.DispCallFunc(ObjPtr(m_enum), 6 * DispCall.POINTER_LENGTH, _
                                               DispCall.CALLCONV.CC_STDCALL, vbLong, 1, _
                                               vtClone, pvClone, result)

    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.Clone", _
            "IEnumVARIANT::Clone returned error"
        Exit Function
    End If

    If ppEnum Is Nothing Then
        Err.Raise ErrorCodes.InvalidProcedureCall, "IEnumVARIANTWrapper.Clone", _
            "IEnumVARIANT::Clone returned nothing"
        Exit Function
    End If

    Dim wrapper As IEnumVARIANTWrapper: Set wrapper = New IEnumVARIANTWrapper
    wrapper.InitFromEnumerator ppEnum
    Set Clone = wrapper
End Function