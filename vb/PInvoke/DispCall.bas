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

' --- struct _complex for C math functions ---
Public Type Complex
    x As Double
    y As Double
End Type

' --- Pointer size for vtable dispatch (4 on 32-bit, 8 on 64-bit) ---
#If Win64 Then
    Public Const POINTER_LENGTH As Long = 8
#Else
    Public Const POINTER_LENGTH As Long = 4
#End If

' --- VARTYPE constants ---
Public Const VT_BYREF As Integer = &H4000   ' only VARTYPE not built into VB6

' --- CALLCONV (oaidl.h) — calling convention for DispCallFunc ---
Public Enum CALLCONV
    CC_FASTCALL = 0
    CC_CDECL = 1
    CC_MSCPASCAL = 2
    CC_PASCAL = 2           ' = CC_MSCPASCAL
    CC_MACPASCAL = 3
    CC_STDCALL = 4
    CC_FPFASTCALL = 5
    CC_SYSCALL = 6
    CC_MPWCDECL = 7
    CC_MPWPASCAL = 8
    CC_MAX = 9
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

' --- Dynamic library loading for CDeclCall ---
Public Declare Function LoadLibraryW Lib "kernel32" ( _
    ByVal lpLibFileName As Long) As Long           ' LPCWSTR — pass StrPtr(name)
    
Public Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String) As Long            ' LPCSTR — ANSI function name

Public Declare Function FreeLibrary Lib "kernel32" ( _
    ByVal hModule As Long) As Long

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

    ' GUID_NULL (16 bytes of zeros — riid parameter, static to avoid per-call zero-init)
    Static nullGuid(0 To 3) As Long
    Dim lcid As Long: lcid = 0

    Dim cArgs As Long
    On Error Resume Next
    cArgs = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then cArgs = 0: Err.Clear
    On Error GoTo 0

    ' Build DISPPARAMS from args (reversed order — IDispatch convention)
    Dim params As DISPPARAMS
    If cArgs > 0 Then
        ' Create reversed copy: IDispatch::Invoke expects last argument
        ' first in the rgvarg array (right-to-left calling convention)
        ReDim argCopy(0 To cArgs - 1) As Variant
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
    Dim excepInfo As excepInfo          ' zero-initialised by VB6 Dim
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
    Dim dispHr As Long: dispHr = DispCallFunc(ObjPtr(obj), VTABLE_INVOKE, CALLCONV.CC_STDCALL, vbLong, 8, _
                          vt(0), pv(0), invokeHr)
    If invokeHr = DISP_E_EXCEPTION Then
        CallByDispId = excepInfo.scode
    Else
        CallByDispId = invokeHr
    End If
End Function

' Returns the decorated export name for a CDECL function.
' Win32: prepends "_" (e.g., "_wsprintf")
' Win64: no decoration (name is used as-is)
Public Function CDeclMangle(ByVal funcName As String) As String
#If Win64 Then
    CDeclMangle = funcName
#Else
    CDeclMangle = "_" & funcName
#End If
End Function

' Calls a C function (CDECL calling convention) at the given address.
'
' Parameters:
'   funcPtr    - Address of the function to call (from GetProcAddress, AddressOf, etc.)
'   resultType - VARTYPE of the return value (vbLong, vbString, vbEmpty for void, etc.)
'   resultVar  - Variant that receives the return value
'   args       - Variant array of positional arguments
'
' Returns:
'   HRESULT from DispCallFunc (0 = S_OK).
'   The actual function return value is written into resultVar.
Public Function CDeclCall(ByVal funcPtr As Long, _
                          ByVal resultType As Integer, _
                          ByRef resultVar As Variant, _
                          ByRef args() As Variant) As Long
    Dim cActuals As Long
    On Error Resume Next
    cActuals = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then cActuals = 0: Err.Clear
    On Error GoTo 0
    
    If cActuals > 0 Then
        ' Build prgvt (VARTYPE of each arg) and prgpvarg (VarPtr of each arg)
        Dim prgvt() As Integer
        Dim prgpvarg() As Long
        ReDim prgvt(0 To cActuals - 1)
        ReDim prgpvarg(0 To cActuals - 1)
        
        Dim i As Long
        For i = 0 To cActuals - 1
            prgvt(i) = VarType(args(i))
            prgpvarg(i) = VarPtr(args(i))
        Next i
        
        CDeclCall = DispCallFunc(0, funcPtr, CALLCONV.CC_CDECL, resultType, _
                                 cActuals, prgvt(0), prgpvarg(0), resultVar)
    Else
        ' No arguments: pass dummy variables for ByRef prgvt/prgpvarg
        Dim vt As Integer
        Dim pv As Long
        CDeclCall = DispCallFunc(0, funcPtr, CALLCONV.CC_CDECL, resultType, _
                                 0, vt, pv, resultVar)
    End If
End Function
