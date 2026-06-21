Attribute VB_Name = "DispCall"
Option Explicit

' --- DISPPARAMS structure for IDispatch::Invoke ---
Private Type DISPPARAMS
    rgvarg As Long            ' VARIANTARG* — array of arguments
    rgdispidNamedArgs As Long ' DISPID* — named argument IDs
    cArgs As Long             ' Number of arguments
    cNamedArgs As Long        ' Number of named arguments
End Type

' --- EXCEPINFO for IDispatch::Invoke exception handling ---
Private Type EXCEPINFO
    wCode As Integer          ' WORD  — error code
    wReserved As Integer      ' WORD  — reserved
    bstrSource As Long        ' BSTR  — source (pointer)
    bstrDescription As Long   ' BSTR  — description (pointer)
    bstrHelpFile As Long      ' BSTR  — help file (pointer)
    dwHelpContext As Long     ' DWORD — help context
    pvReserved As Long        ' PVOID — reserved
    pfnDeferredFillIn As Long ' PVOID — deferred fill-in
    scode As Long             ' SCODE — actual error code
End Type

' --- Pointer size for vtable dispatch (4 on 32-bit, 8 on 64-bit) ---
#If Win64 Then
    Public Const POINTER_LENGTH As Long = 8
#Else
    Public Const POINTER_LENGTH As Long = 4
#End If

' --- VARTYPE constants ---
Public Const VT_BYREF As Integer = &H4000   ' only VARTYPE not built into VB6

' --- Calling convention ---
Public Enum CallingConvention
    StdCall = 4
End Enum

' --- IDispatch vtable offsets ---
' IUnknown:    QueryInterface(0), AddRef(1), Release(2)
' IDispatch:   GetTypeInfoCount(3), GetTypeInfo(4), GetIDsOfNames(5), Invoke(6)
Private Const VTABLE_INVOKE As Long = 6 * POINTER_LENGTH

Private Const DISP_E_MEMBERNOTFOUND As Long = -2147352573
Private Const DISP_E_EXCEPTION As Long = -2147352567

Public Declare Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As Long, _
    ByRef pvargResult As Variant) As Long
Public Function TryGetEnum(ByRef source As Object, ByRef enumOut As IUnknown) As Long
    Dim dispObj As IUnknown: Set dispObj = source
    Dim ObjEnum As Variant: ObjEnum = Empty
    Dim callResult As Long: callResult = CallByDispId(dispObj, -4, 2, vbDataObject, ObjEnum)
    If callResult = DISP_E_MEMBERNOTFOUND Then
        callResult = CallByDispId(dispObj, -4, 1, vbDataObject, ObjEnum)
    End If
    If callResult = 0 And VarType(ObjEnum) = vbDataObject Then
        Set enumOut = ObjEnum
    End If
    
    TryGetEnum = callResult
End Function

' Calls IDispatch::Invoke on obj using the given DISPID directly,
' bypassing GetIDsOfNames name resolution. This works for hidden/
' restricted members (_NewEnum, etc.) that CallByName cannot access.
'
' Parameters:
'   obj      - The IDispatch object to call on
'   dispId   - The DISPID to invoke (e.g., -4 for _NewEnum)
'   callType - VBA call type: 1=VbMethod(DISPATCH_METHOD),
'              2=VbGet(DISPATCH_PROPERTYGET), 4=VbLet(DISPATCH_PROPERTYPUT),
'              8=VbSet(DISPATCH_PROPERTYPUTREF)
'   resultType - the return value's variant type
'   resultVar - the variant that will hold the return value
'   args     - Optional positional arguments passed to the method/property
'
' Returns:
'   The value/object returned by Invoke, or Empty on failure (check VarType).
Public Function CallByDispId(ByRef obj As Object, _
                             ByVal dispId As Long, _
                             ByVal callType As Integer, _
                             ByVal resultType As Integer, _
                             ByRef resultVar As Variant, _
                             ParamArray args() As Variant) As Long

    ' GUID_NULL (16 bytes of zeros — riid parameter)
    Dim nullGuid(0 To 3) As Long
    Dim lcid As Long: lcid = 0

    ' Build DISPPARAMS from args (reversed order — IDispatch convention)
    Dim params As DISPPARAMS

    Dim cArgs As Long
    On Error Resume Next
    cArgs = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then cArgs = 0: Err.Clear
    On Error GoTo 0

    If cArgs > 0 Then
        ' Create reversed copy: IDispatch::Invoke expects last argument
        ' first in the rgvarg array (right-to-left calling convention)
        Dim argCopy() As Variant
        ReDim argCopy(0 To cArgs - 1)
        Dim i As Long
        For i = 0 To cArgs - 1
            argCopy(i) = args(cArgs - 1 - i)
        Next i

        params.rgvarg = VarPtr(argCopy(0))
        params.cArgs = cArgs
    Else
        params.rgvarg = 0
        params.cArgs = 0
    End If

    params.cNamedArgs = 0
    params.rgdispidNamedArgs = 0

    ' Parameter VARTYPE descriptors (left-to-right, matches Invoke signature)
    Dim vt(0 To 7) As Integer
    vt(0) = vbLong                     ' dispIdMember
    vt(1) = vbLong Or VT_BYREF         ' riid (pointer to GUID_NULL)
    vt(2) = vbLong                     ' lcid
    vt(3) = vbInteger                  ' wFlags
    vt(4) = vbLong Or VT_BYREF         ' pDispParams (pointer to DISPPARAMS)
    vt(5) = resultType Or VT_BYREF     ' pVarResult (pointer to Variant)
    vt(6) = vbLong Or VT_BYREF         ' pExcepInfo (pointer to EXCEPINFO)
    vt(7) = vbLong Or VT_BYREF         ' puArgErr (pointer to argErr)

    ' Wrap each argument in a Variant — prgpvarg[i] always points to one
    Dim vDispId As Variant: vDispId = CLng(dispId)
    Dim vRiid As Variant: vRiid = VarPtr(nullGuid(0))
    Dim vLcid As Variant: vLcid = lcid
    Dim vWFlags As Variant: vWFlags = callType
    Dim vDispParams As Variant: vDispParams = VarPtr(params)
    Dim pVarResult As Variant: pVarResult = VarPtr(resultVar)
    Dim excepInfo As EXCEPINFO          ' zero-initialised by VB6 Dim
    Dim vExcepInfo As Variant: vExcepInfo = CLng(VarPtr(excepInfo))
    Dim argErr As Long: argErr = -1
    Dim vArgErr As Variant: vArgErr = VarPtr(argErr)

    ' Parameter value pointers — all point to Variant wrappers
    Dim pv(0 To 7) As Long
    pv(0) = VarPtr(vDispId)          ' Variant containing -4
    pv(1) = VarPtr(vRiid)            ' Variant containing GUID_NULL pointer
    pv(2) = VarPtr(vLcid)            ' Variant containing 0
    pv(3) = VarPtr(vWFlags)          ' Variant containing callType (2)
    pv(4) = VarPtr(vDispParams)      ' Variant containing DISPPARAMS pointer
    pv(5) = VarPtr(pVarResult)       ' Variant containing pVarResult pointer
    pv(6) = VarPtr(vExcepInfo)       ' Variant containing pointer to EXCEPINFO
    pv(7) = VarPtr(vArgErr)          ' Variant containing pointer to argErr

    Dim invokeHr As Variant

    Dim dispHr As Long
    dispHr = DispCallFunc(ObjPtr(obj), VTABLE_INVOKE, CallingConvention.StdCall, vbLong, 8, _
                          vt(0), pv(0), invokeHr)
    If invokeHr = DISP_E_EXCEPTION Then
        CallByDispId = excepInfo.scode
    Else
        CallByDispId = invokeHr
    End If
End Function
