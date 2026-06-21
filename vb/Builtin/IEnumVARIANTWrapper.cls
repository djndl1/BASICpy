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
    Dim callResult As Long
    callResult = DispCall.TryGetEnum(source, enumObj)
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
    Dim vCelt As Variant: vCelt = CLng(count)
    Dim pvBuffer As Variant: pvBuffer = VarPtr(buffer(0))
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
        DispCall.CallingConvention.StdCall, vbLong, 3, _
        vt(0), pv(0), result)

    ' COM failure should propagate as a VB6 error
    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.NextItems", _
            "IEnumVARIANT::Next returned error"
        Exit Function
    End If

    ' S_OK (0 = CommonHResults.Success) with fetched > 0 → items available
    If VarType(result) = vbLong Then
        If CLng(result) = CommonHResults.Success And vFetched > 0 Then
            If vFetched = count Then
                NextItems = buffer
            Else
                ' S_FALSE means partial fetch — slice to actual count
                Dim partial() As Variant
                ReDim partial(0 To vFetched - 1)
                Dim i As Long
                For i = 0 To vFetched - 1
                    partial(i) = buffer(i)
                Next i
                NextItems = partial
            End If
            Exit Function
        End If
    End If

    ' S_FALSE or empty fetch → exhausted, return empty array
    NextItems = EmptyVariantArray
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
    Dim pv(0 To 0) As Long: pv(0) = VarPtr(vCelt)
    Dim vt(0 To 0) As Integer: vt(0) = vbLong
    Dim result As Variant
    Dim hr As Long
    hr = DispCall.DispCallFunc(ObjPtr(m_enum), 4 * DispCall.POINTER_LENGTH, _
        DispCall.CallingConvention.StdCall, vbLong, 1, _
        vt(0), pv(0), result)

    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.Skip", _
            "IEnumVARIANT::Skip returned error"
    End If
End Sub

' Resets the enumerator back to the beginning.
' After Reset, the next NextItems call returns items from the start.
Public Sub Reset()
    ReDim vt(0) As Integer
    ReDim pv(0) As Long
    Dim result As Variant
    Dim hr As Long
    hr = DispCall.DispCallFunc(ObjPtr(m_enum), 5 * DispCall.POINTER_LENGTH, _
        DispCall.CallingConvention.StdCall, vbLong, 0, _
        vt(0), pv(0), result)

    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.Reset", _
            "IEnumVARIANT::Reset returned error"
    End If
End Sub
Public Function Clone() As IEnumVARIANTWrapper
    Dim ppEnum As IUnknown
    Dim vtClone(0) As Integer
    Dim pvClone(0) As Long
    Dim result As Variant
    Dim hr As Long

    vtClone(0) = vbDataObject Or DispCall.VT_BYREF
    pvClone(0) = VarPtr(ppEnum)

    hr = DispCall.DispCallFunc(ObjPtr(m_enum), 6 * DispCall.POINTER_LENGTH, _
        DispCall.CallingConvention.StdCall, vbLong, 1, _
        vtClone(0), pvClone(0), result)

    If hr <> 0 Then
        Err.Raise hr, "IEnumVARIANTWrapper.Clone", _
            "IEnumVARIANT::Clone returned error"
        Exit Function
    End If

    If ppEnum Is Nothing Then
        Err.Raise 5, "IEnumVARIANTWrapper.Clone", _
            "IEnumVARIANT::Clone returned nothing"
        Exit Function
    End If

    Dim wrapper As IEnumVARIANTWrapper
    Set wrapper = New IEnumVARIANTWrapper
    wrapper.InitFromEnumerator ppEnum
    Set Clone = wrapper
End Function
