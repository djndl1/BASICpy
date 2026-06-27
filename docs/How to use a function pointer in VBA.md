VB6 and VBA come with no support for function pointers.

Also, when you wish to execute a function in a dll using the Declare function, you can only call functions created by the Steadcall calling conversation.

These constraints can be avoided by using the DispCallFunc API.
The DispCallFunc is widely used in VB6 when erasing the history of IE.
Although the DispCallFunc is known as API for calling the IUnknown interface, in fact, you can also perform other functions other than COM by passing the NULL to the first argument.

As explained in the http://msdn.microsoft.com/en-us/library/ms221473(v=vs.85).aspx , the DispCallFunc argument is as follows.

```c
HRESULT DispCallFunc(
  void *pvInstance,
  ULONG_PTR oVft,
  CALLCONV cc,
  VARTYPE vtReturn,
  UINT cActuals,
  VARTYPE *prgvt,
  VARIANTARG **prgpvarg,
  VARIANT *pvargResult
);
```

Use of the argument sixth and seventh is very peculiar, and there are few descriptions about it available. Therefore, I'll try to summarize it this time around.

First argument, pvInstance
 In this time, I am going to call functions other than COM. Thus, always setting at 0.

Second argument, oVft
 Passing the address of a function.

Third argument, cc
 Passing the function's calling convention.
 For `CALLCONV` type, the following values are defined for the OAIdl.h.
 
 ```c
 enum tagCALLCONV
    {	CC_FASTCALL	= 0,
	CC_CDECL	= 1,
	CC_MSCPASCAL	= CC_CDECL + 1,
	CC_PASCAL	= CC_MSCPASCAL,
	CC_MACPASCAL	= CC_PASCAL + 1,
	CC_STDCALL	= CC_MACPASCAL + 1,
	CC_FPFASTCALL	= CC_STDCALL + 1,
	CC_SYSCALL	= CC_FPFASTCALL + 1,
	CC_MPWCDECL	= CC_SYSCALL + 1,
	CC_MPWPASCAL	= CC_MPWCDECL + 1,
	CC_MAX	= CC_MPWPASCAL + 1
    } 	CALLCONV;
```

  For a general WIN32API function, `CC_STDCALL` (= 4) is used.
  For a function that is specifically designed for language C, `CC_CDECL` (= 1) is used. 

Fourth argument, vtReturn
 Setting the value of `VbVarType` enumerated type that indicates the type of the function's return value.
 In the event where the function does not return a value (the Sub procedure referred to in VB), setting the vbEmpty.
 VARTYPE type as defined in WTypes.h is shown below.
 typedef unsigned short VARTYPE;
 By ignoring the existence of the code, it is equivalent to the Integer type in VB.
 
 
Fifth argument, cActuals
 Setting the number of function arguments.

Sixth argument, prgvt
 Storing in an array of Integer type VbVarType enumerated values that indicate the type of the function's arguments, and then passing the address of the beginning of the array.
 However, when the number of the arguments is 0, passing 0.

Seventh argument, prgpvarg
 Once again, storing each function's argument in Variant type, and storing the address of each of the variant variable in the Long type array that is prepared separately, then passing the address of the beginning of the array.
 However, when the number of arguments is 0, then passing 0.

Then, the return value of the DispCallFunc API is, if a call to the function is done successfully in the Long type, S_OK (= 0) is returned.

As an example, the API declaration and the declaration of the enumeration in VB are as follows.

```vb
Private Declare Function DispCallFunc Lib "OleAut32.dll" _
(ByVal pvInstance As Long, _
   ByVal oVft As Long, _
   ByVal cc As Long, _
   ByVal vtReturn As Integer, _
   ByVal cActuals As Long, _
   ByVal prgvt As Long, _
   ByVal prgpvarg As Long, _
   ByVal pvargResult As Long) As Long

Enum tagCALLCONV
    CC_FASTCALL = 0
    CC_CDECL = 1
    CC_MSCPASCAL = CC_CDECL + 1
    CC_PASCAL = CC_MSCPASCAL
    CC_MACPASCAL = CC_PASCAL + 1
    CC_STDCALL = CC_MACPASCAL + 1
    CC_FPFASTCALL = CC_STDCALL + 1
    CC_SYSCALL = CC_FPFASTCALL + 1
    CC_MPWCDECL = CC_SYSCALL + 1
    CC_MPWPASCAL = CC_MPWCDECL + 1
    CC_MAX = CC_MPWPASCAL
End Enum
```

First, I tried the simplest way with a pointer to a function with no defined arguments in a standard module of VBA.
In VBA, as long as functions that are defined in a standard module, you can obtain those addresses using AddressOf operator.
Click here to download the following code as a text file.

```vb
'You can use a sample code for this site freely.
'Though this is not a duty, I am grateful that you describe that you reffered this site(http://akihitoyamashiro.com/en/VBA/), 
'when you present this sample code in your web site.

Public Sub NoParamNoReturn()
    MsgBox "NoParamNoReturn is called."
End Sub

Public Sub Test1()
    Dim lDispCallFuncResult As Long
    
    'Declare a Variant variable that will store the return value of the function called.
    'This variable is needed even if the function does not return a value.
    Dim vFuncResult As Variant
    
    'In this case, there is no need to initialize the variable after declaring it.
    'However, when using a Variant variable, it can be initialized like this
    'by setting it to Empty.
    vFuncResult = Empty
    
    'You can get the address of a function in VBA by using the AddressOf operator.
    'Enter the name of the enumerated type; at the point when you type in the period you will be able to use VBA's input support functionality.
    'This time we are calling a sub procedure so there is no value to be returned.
    'Hence we set the data type of the return variable to be vbEmpty(=0).
    lDispCallFuncResult = DispCallFunc(0, AddressOf NoParamNoReturn, _
                                       tagCALLCONV.CC_STDCALL, VbVarType.vbEmpty, _
                                       0, 0, 0, VarPtr(vFuncResult))
    
    'When the call to the DispCallFunc API is successful, it returns 0.
    'However, this only means that the call to the DispCallFunc API was successful,
    'and it does not indicate whether the target function call was successful or not.
    Debug.Print lDispCallFuncResult
    
End Sub
```

The next section we also try using pointers of functions defined in VBA's standard module.
However, in the next function there are variables and return values.
Click here to download the following code as a text file.

```vb
Option Explicit

'You can use a sample code for this site freely.
'Though this is not a duty, I am grateful that you describe that you reffered this site(http://akihitoyamashiro.com/en/VBA/), 
'when you present this sample code in your web site.

Private Declare Function DispCallFunc Lib "OleAut32.dll" _
(ByVal pvInstance As Long, _
   ByVal oVft As Long, _
   ByVal cc As Long, _
   ByVal vtReturn As Integer, _
   ByVal cActuals As Long, _
   ByVal prgvt As Long, _
   ByVal prgpvarg As Long, _
   ByVal pvargResult As Long) As Long

Enum tagCALLCONV
    CC_FASTCALL = 0
    CC_CDECL = 1
    CC_MSCPASCAL = CC_CDECL + 1
    CC_PASCAL = CC_MSCPASCAL
    CC_MACPASCAL = CC_PASCAL + 1
    CC_STDCALL = CC_MACPASCAL + 1
    CC_FPFASTCALL = CC_STDCALL + 1
    CC_SYSCALL = CC_FPFASTCALL + 1
    CC_MPWCDECL = CC_SYSCALL + 1
    CC_MPWPASCAL = CC_MPWCDECL + 1
    CC_MAX = CC_MPWPASCAL
End Enum

Public Function SixParamOneReturn( _
                ByVal longVal As Long, ByRef longRef As Long, _
                ByVal byteVal As Byte, ByRef byteRef As Byte, _
                ByVal strVal As String, ByRef strRef As String _
                ) As Long
                
    'ByVal...String in VB is equivalent to BSTR in VC++.
    'ByRef...String in VB is equivalent to BSTR* in VC++. 
                
    Debug.Print "SixParamOneReturn"
    Debug.Print "longVal = " & longVal
    Debug.Print "longRef = " & longRef
    Debug.Print "byteVal = " & byteVal
    Debug.Print "byteRef = " & byteRef
    Debug.Print "strVal = " & strVal
    Debug.Print "strRef = " & strRef
    Debug.Print
    
    SixParamOneReturn = longVal + longRef
    
    longVal = 100
    longRef = 200
    byteVal = 10
    byteRef = 20
    strVal = "strVal"
    strRef = "strRef"
    MsgBox "EightParamOneReturn is called."
End Function

Public Sub Test2()
    'The annotations that were explained in Test1 have been omitted.

    Dim lDispCallFuncResult As Long    
    Dim vFuncResult As Variant

    vFuncResult = Empty
    
    Dim lVal As Long, vlVal As Variant
    Dim lRef As Long, vlRef As Variant
    Dim bVal As Byte, vbVal As Variant
    Dim bRef As Byte, vbRef As Variant
    Dim sVal As String, vsVal As Variant
    Dim sRef As String, vsRef As Variant
    
    lVal = 12
    lRef = 34
    bVal = 56
    bRef = 78
    sVal = "string1"
    sRef = "string2"
    
    'DispCallFunc method
    'The variables passed to the called function all have to 
    'be passed in Variant form so all the variables are 
    'declared as Variant variables.
    
    'It is easier to create a Variant array and assign the 
    'array but to show there is no need for it to be an 
    'array, we will not create an array.

    'Calling by reference, as instructed by VB's ByRef, 
    'involves not declaring variables, but we have to assign 
    'pointers to variables.
        
    vlVal = lVal 'Because it is the ByVal variable, set a value.
    vlRef = VarPtr(lRef) 'It is the ByVal variable so assign 
                         'pointers. In C it corresponds to &IRef.
    vbVal = bVal
    vbRef = VarPtr(bRef)
    vsVal = sVal
    vsRef = VarPtr(sRef)
    
    'After declaring the Variant variables
    'along with set each Variant variable's type as an integer array.
    'we set each Variant variable's address as a long array.
    
    Dim iVarTypes(0 To 5) As Integer 'hard to tell, but 
                                     'first letter is a small letter i.
    Dim lVarPtrs(0 To 5) As Long
    
    iVarTypes(0) = VarType(vlVal) 'VbVarType.vbLong is assigned
    lVarPtrs(0) = VarPtr(vlVal)   '"VarPtr(vlVal)" corresponds to 
                                  '"&vlVal" in C.
    
    iVarTypes(1) = VarType(vlRef) 'The pointer is in Long form
                                  'so VbVarType.vbLong is assigned.
    lVarPtrs(1) = VarPtr(vlRef)
    
    iVarTypes(2) = VarType(vbVal) 'VbVarType.vbByte is assigned.
    lVarPtrs(2) = VarPtr(vbVal)
    
    iVarTypes(3) = VarType(vbRef) 'The pointer is in Long form 
                                  'so VbVarType.vbLong is assigned.
    lVarPtrs(3) = VarPtr(vbRef)
    
    iVarTypes(4) = VarType(vsVal) 'VbVarType.vbString is assigned.
    lVarPtrs(4) = VarPtr(vsVal)
    
    iVarTypes(5) = VarType(vsRef) 'The pointer is in Long form
                                  'so VbVarType.vbLong is assigned.
    lVarPtrs(5) = VarPtr(vsRef)
        
    'The returned value's type is long so we set the 4th argument 
    '(the returned value's type) as vbLong(=3).
    
    'There are 6 variables, we set the 5th argument to 6.

    'For the 6th argument we set the first address in the 
    'integer array that stores each variant variable's inner type.

    'For the 7th argument we set the first address in the 
    'long array that stores each variant variable's address.

    lDispCallFuncResult = DispCallFunc(   _
                                       0, _
                                       AddressOf SixParamOneReturn, _
                                       tagCALLCONV.CC_STDCALL, _
                                       VbVarType.vbLong, _
                                       6, _
                                       VarPtr(iVarTypes(0)), _
                                       VarPtr(lVarPtrs(0)), _
                                       VarPtr(vFuncResult) _
                                       )

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'In this section, if the declaration of the DispCallFunc 
    'is as follows, then the 6th, 7th and 8th arguments VarPtr are 
    'unnecessary.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Private Declare Function DispCallFunc Lib "OleAut32.dll" _
    (ByVal pvInstance As Long, _
       ByVal oVft As Long, _
       ByVal cc As Long, _
       ByVal vtReturn As Integer, _
       ByVal cActuals As Long, _
       ByRef prgvt As Integer, _
       ByRef prgpvarg As Long, _
       ByRef pvargResult As Variant) As Long


    'Regarding the 6th, 7th and 8th arguments, we change 
    'ByVal pointer's type to ByRef variable's type

    '6th argument ByVal ... Long to ByRef ... Integer
    '7th argument ByVal ... Long to ByRef ... Long
    '8th argument ByVal ... Long to ByRef ... Variant

    'Because of this getting the address value is done 
    'automatically by VBA, and we can do the DispCallFunc 
    'part in the following way

    
    
    'lDispCallFuncResult = DispCallFunc( _
                                        0, _
                                        AddressOf SixParamOneReturn, _
                                        tagCALLCONV.CC_STDCALL, _
                                        VbVarType.vbLong, _
                                        6, _
                                        iVarTypes(0), _
                                        lVarPtrs(0), _
                                        vFuncResult _
                                        )
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
    Debug.Print "Test2"
    Debug.Print "lDispCallFuncResult = " & lDispCallFuncResult
    
    Debug.Print "vFuncResult = " & vFuncResult
    
    Debug.Print "lVal = " & lVal 'lVal value is left as is.
    Debug.Print "lRef = " & lRef 'lRef value is passed to a pointer 
                                 'so the original value is affected 
                                 'by the changes where it is called.

    Debug.Print "bVal = " & bVal
    Debug.Print "bRef = " & bRef
    Debug.Print "sVal = " & sVal
    Debug.Print "sRef = " & sRef
    
End Sub
```

Until now, I assumed you understand DispCallFunc's call procedure. Actually, there are not many cases where functions within VBA is called with DispCallFunc.
Also, because standard WIN32API are created with the STDCALL calling convention, you can call them if you do a Declare Function. In such a case, there are few cases where calls are made with DispCallFunc.

Thus I will try to call functions created with the CDECL calling convention.
However, to do this requires some work to create DLLs, so I will try to call the wsprintf function, which is an WIN32API created with the CDECL calling convention, unlike other WIN32API.
This function has variable arity, so it is created with CDECL. Therefore, it cannot be called in VBA by doing a Declare Function.

The wsprintf function has wsprintfW function for Unicode and wsprintfA function for ANSI. I will try using the wsprintfW function, which is highly compatible with VBA's string format.

Click here to download the following code as a text file.

```vb
Option Explicit


'You can use a sample code for this site freely.
'Though this is not a duty, I am grateful that you describe that you reffered this site(http://akihitoyamashiro.com/en/VBA/), 
'when you present this sample code in your web site.

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" _
    (ByVal lpFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32.dll" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32.dll" _
    (ByVal hModule As Long) As Long

Private Declare Function DispCallFunc Lib "OleAut32.dll" _
    (ByVal pvInstance As Long, _
        ByVal oVft As Long, _
        ByVal cc As Long, _
        ByVal vtReturn As Integer, _
        ByVal cActuals As Long, _
        ByVal prgvt As Long, _
        ByVal prgpvarg As Long, _
        ByVal pvargResult As Long) As Long
   
Enum tagCALLCONV
    CC_FASTCALL = 0
    CC_CDECL = 1
    CC_MSCPASCAL = CC_CDECL + 1
    CC_PASCAL = CC_MSCPASCAL
    CC_MACPASCAL = CC_PASCAL + 1
    CC_STDCALL = CC_MACPASCAL + 1
    CC_FPFASTCALL = CC_STDCALL + 1
    CC_SYSCALL = CC_FPFASTCALL + 1
    CC_MPWCDECL = CC_SYSCALL + 1
    CC_MPWPASCAL = CC_MPWCDECL + 1
    CC_MAX = CC_MPWPASCAL
End Enum


Public Sub Test3()
    'The annotations that were explained in Test1,Test2 have been omitted.
    
    Dim lDispCallFuncResult As Long
    Dim vFuncResult As Variant

    vFuncResult = Empty
    
    Dim lLibraryHandle As Long
    lLibraryHandle = LoadLibrary("user32.dll")
    If lLibraryHandle = 0 Then
            Debug.Print "LoadLibrary failed."
            Exit Sub
    End If
    
    Dim lProcAddress As Long
    lProcAddress = GetProcAddress(lLibraryHandle, "wsprintfW")
    If lProcAddress = 0 Then
            Debug.Print "GetProcAddress failed."
            FreeLibrary lLibraryHandle
            Exit Sub
    End If

    'wsprintfW has the following arguments.
    'LPTSTR lpOut , LPCTSTR lpFmt , ...
    
    'sprintf works almost exactly the same as sprintf in the C language.
    
    'First argument: buffer that will store the formatted output
    'Second argument: string specifying desired format
    'Third argument (and additional arguments): strings and values you want to embed
    
    'This time we will embed the phrase "Param1 = %d , Param2 = %s" with a number (%d) and a string (%s).
    
    Dim sOut As String
    Dim sFormat As String
    Dim lParam1 As Long
    Dim sParam2 As String
    
    sOut = String(200, vbNullChar) 'Buffer for Unicode 200 characters is secured, embedded using \0(=400 bytes).
    sFormat = "Param1 = %d , Param2 = %s"
    lParam1 = 123456
    sParam2 = "abc"
    
    'This time, 4 arguments are passed to wsprintfW, so 4 Variant format variables are prepared.
    'This time, preparing 4 separate variables takes work, so an array is prepared.
    Dim vParams(0 To 3) As Variant
        
    'LP(C)TSTR for wsprintfW is ultimately wchar_t*, so we must pass VBA's StrPtr(String type variable).
    vParams(0) = StrPtr(sOut)
    vParams(1) = StrPtr(sFormat)
    vParams(2) = lParam1
    vParams(3) = StrPtr(sParam2)
    
    'This time, we consider the appropriation by other code, and declare the following array variables dynamically.
    Dim iVarTypes() As Integer
    Dim lVarPtrs() As Long
    
    ReDim iVarTypes(0 To UBound(vParams))
    ReDim lVarPtrs(0 To UBound(vParams))

    
    Dim i As Long
    For i = 0 To UBound(vParams)
        iVarTypes(i) = VarType(vParams(i))
        lVarPtrs(i) = VarPtr(vParams(i))
    Next
    
    'We specify the return value of GetProcAdress for the 2nd argument.
    
    'wsprintfW function's calling convention is CDECL, so we specify CC_CDECL for the 3rd argument.
    
    'This time, the return value is int type under VC++ ( = Long type under VB), so
    'we specify the return value's type as vbLong(=3).
    
    'There are 4 arguments. For portability reason, without directly specifying 4, 
    'the upper limit of the array is set using Ubound(vParams)+1.
    
    lDispCallFuncResult = DispCallFunc(0, lProcAddress, _
                                       tagCALLCONV.CC_CDECL, VbVarType.vbLong, _
                                        UBound(vParams) + 1, VarPtr(iVarTypes(0)), VarPtr(lVarPtrs(0)), VarPtr(vFuncResult))

    
    FreeLibrary lLibraryHandle
    
    Debug.Print "Test3"
    Debug.Print "lDispCallFuncResult = " & lDispCallFuncResult
    
    Debug.Print "vFuncResult = " & vFuncResult
    
    'sOut's Unicode string buffer address is specified as the first argument
    'so the wsprintfW API overwrites the beginning part of sOut which is padded with 200 ChrW(0) characters 
    'with a null terminated character string. 
    'Therefore, the character string wanted by VB is the left-hand string from the first seen ChrW(0)(=vbNullChar).

    Debug.Print Left(sOut, InStr(1, sOut, vbNullChar) - 1)

    'However, even without using InStr, wsprintfW returns the length of the formatted character string, so 
    'Debug.Print Left(sOut, vFuncResult)
    ' is what we can do.


End Sub
```

Next, let's take a look at the `wsprintfA` function, which has low compatibility with VBA's string format.
Because the entire character string buffer needs to be ANSI character string (same as SJIS character string), explicit reciprocal conversion between Unicode character string and ANSI character string in the program code is required to process string format that has Unicode character string buffer.

The is really a different topic from DispCallFunc. But because I expect that you will be often be dealing with legacy modules using the ANSI character string system in cases of using the CDECL call system, I bring this up here.
Please note that when VBA's Declare Function is As String, this conversion takes place automatically.

Click here to download the following code as a text file.

```vb
Option Explicit


'You can use a sample code for this site freely.
'Though this is not a duty, I am grateful that you describe that you reffered this site(http://akihitoyamashiro.com/en/VBA/), 
'when you present this sample code in your web site.

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" _
    (ByVal lpFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32.dll" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32.dll" _
    (ByVal hModule As Long) As Long

Private Declare Function DispCallFunc Lib "OleAut32.dll" _
    (ByVal pvInstance As Long, _
        ByVal oVft As Long, _
        ByVal cc As Long, _
        ByVal vtReturn As Integer, _
        ByVal cActuals As Long, _
        ByVal prgvt As Long, _
        ByVal prgpvarg As Long, _
        ByVal pvargResult As Long) As Long
   
Enum tagCALLCONV
    CC_FASTCALL = 0
    CC_CDECL = 1
    CC_MSCPASCAL = CC_CDECL + 1
    CC_PASCAL = CC_MSCPASCAL
    CC_MACPASCAL = CC_PASCAL + 1
    CC_STDCALL = CC_MACPASCAL + 1
    CC_FPFASTCALL = CC_STDCALL + 1
    CC_SYSCALL = CC_FPFASTCALL + 1
    CC_MPWCDECL = CC_SYSCALL + 1
    CC_MPWPASCAL = CC_MPWCDECL + 1
    CC_MAX = CC_MPWPASCAL
End Enum


Public Sub Test4()
    'The annotations that were explained in Test1,Test2,Test3 have been omitted.
    
    Dim lDispCallFuncResult As Long
    Dim vFuncResult As Variant

    vFuncResult = Empty
    
    Dim lLibraryHandle As Long
    lLibraryHandle = LoadLibrary("user32.dll")
    If lLibraryHandle = 0 Then
            Debug.Print "LoadLibrary failed."
            Exit Sub
    End If    

    Dim lProcAddress As Long
    'Unlike Test3, lProcAddress holds a pointer to the function wsprintfA.
    lProcAddress = GetProcAddress(lLibraryHandle, "wsprintfA")
    If lProcAddress = 0 Then
            Debug.Print "GetProcAddress failed."
            FreeLibrary lLibraryHandle
            Exit Sub
    End If

    'wsprintfA has the following arguments.
    'LPTSTR lpOut , LPCTSTR lpFmt , ...
    
    'sprintf works almost exactly the same as sprintf in the C language.
    
    'First argument: buffer that will store the formatted output
    'Second argument: string specifying desired format
    'Third argument (and additional arguments): strings and values you want to embed
    
    'This time we will embed the phrase "Param1 = %d , Param2 = %s" with a number (%d) and a string (%s).
    
    Dim sOut As String
    Dim bOut() As Byte     'holds the result from the conversion of the Unicode character string 
                           'in the string type variable sOut to an ANSI character string.
    Dim sFormat As String
    Dim bFormat() As Byte  'holds the result from the conversion of the Unicode character string 
                           'in the string type variable sFormat to an ANSI character string.
    Dim lParam1 As Long
    Dim sParam2 As String
    Dim bParam2() As Byte  'holds the result from the conversion of the Unicode character string 
                           'in the string type variable sParam2 to an ANSI character string.
    
    sOut = String(200, vbNullChar) 'Buffer for Unicode 200 characters is secured, embedded using \0(=400 bytes).
    bOut = StrConv(sOut, vbFromUnicode)  'Convert the Unicode character string to an ANSI character Byte array.

    sFormat = "Param1 = %d , Param2 = %s"
    bFormat = StrConv(sFormat, vbFromUnicode) 'Convert the Unicode character string to an ANSI character Byte array.
                                              'vbFromUnicode means a conversion from Unicode to ANSI.    
    lParam1 = 123456
    
    sParam2 = "abc"
    bParam2 = StrConv(sParam2, vbFromUnicode) 'Convert the Unicode character string to an ANSI character Byte array.
            
    'This time, 4 arguments are passed to wsprintfA, so 4 Variant format variables are prepared.
    'This time, preparing 4 separate variables takes work, so an array is prepared.
    Dim vParams(0 To 3) As Variant
    
    'Because LP(C)TSTR in wsprintfA is ultimately a char*, it is necessary to convert the string type variable to an ANSI character string,
    'store it in a Byte type array, and then pass the beginning address of that Byte type array.    
        
    vParams(0) = VarPtr(bOut(0))
    vParams(1) = VarPtr(bFormat(0))
    vParams(2) = lParam1
    vParams(3) = VarPtr(bParam2(0))
    
    Dim iVarTypes() As Integer
    Dim lVarPtrs() As Long
    
    ReDim iVarTypes(0 To UBound(vParams))
    ReDim lVarPtrs(0 To UBound(vParams))

    Dim i As Long
    For i = 0 To UBound(vParams)
        iVarTypes(i) = VarType(vParams(i))
        lVarPtrs(i) = VarPtr(vParams(i))
    Next
    
    'We specify the return value of GetProcAdress for the 2nd argument.
    
    'wsprintfA function's calling convention is CDECL, so we specify CC_CDECL for the 3rd argument.
    
    'This time, the return value is int type under VC++ ( = Long type under VB), so
    'we specify the return value's type as vbLong(=3).
        
    lDispCallFuncResult = DispCallFunc(0, lProcAddress, _
                                       tagCALLCONV.CC_CDECL, VbVarType.vbLong, _
                                        UBound(vParams) + 1, VarPtr(iVarTypes(0)), VarPtr(lVarPtrs(0)), VarPtr(vFuncResult))

    
    FreeLibrary lLibraryHandle
    
    Debug.Print "Test4"
    Debug.Print "lDispCallFuncResult = " & lDispCallFuncResult
    
    Debug.Print "vFuncResult = " & vFuncResult
    
    'Because the address of bOut's ANSI character string buffer was specified as the first argument,
    'wsprintfA API overwrites the beginning part of bOut which is padded with 200 ChrA(0) characters with a null-terminated string.
    'Therefore, it is necessary here to convert bOut from ANSI to Unicode to make it possible to handle it in VBA.
    'Further, in VB the left side character string is more desirable than the ChrW(0)(=vbNullChar) that was found first.

    sOut = StrConv(bOut, vbUnicode) 'vbUnicode means a conversion from ANSI to Unicode.
    Debug.Print Left(sOut, InStr(1, sOut, vbNullChar) - 1)
End Sub
```