Attribute VB_Name = "InCoreImpl"
Option Explicit

' Core implementation for PyIn — returns True if item is in container.
Public Function PyInCore(ByRef item As Variant, ByRef container As Variant) As Boolean
    ' 0. Undimensioned array container → empty, nothing to find
    If IsArray(container) Then
        If PyLenImpl(container) = -1 Then
            PyInCore = False
            Exit Function
        End If
    End If
    
    ' 1. IContains protocol (constant-time when available)
    If IsObject(container) Then
        If TypeOf container Is IContains Then
            Dim c As IContains: Set c = container
            PyInCore = c.Contains(item)
            Exit Function
        End If
    End If
    
    ' 1b. String container → substring matching (Python: y.find(x) != -1)
    If VarType(container) = vbString Then
        If VarType(item) <> vbString Then
            Ensure.IsTrue False, CommonHResults.TypeMismatch, _
                "Builtin.PyIn", "'in <string>' requires string as left operand"
        End If
        PyInCore = (InStr(1, container, item, vbBinaryCompare) > 0)
        Exit Function
    End If
    
    ' 2. Type-name fast path — known COM containers with O(1) lookup
    If IsObject(container) Then
        Dim dispObj As Object: Set dispObj = container
        Dim matched As Boolean: matched = False
        PyInCore = PyInByTypeName(item, dispObj, matched)
        If matched Then Exit Function
    End If
    
    ' 3. Iteration fallback
    Dim element As Variant: Do While Builtin.PyTryNext(Builtin.PyIter(container), element)
        If element = item Then
            PyInCore = True
            Exit Function
        End If
    Loop
    
    ' 4. Dynamic contains check — last resort for unknown types
    If IsObject(container) Then
        Dim dynResult As Variant: dynResult = PyInDynamicCheck(item, dispObj)
        If Not IsEmpty(dynResult) Then
            PyInCore = dynResult
            Exit Function
        End If
    End If
    
    PyInCore = False
End Function
