Attribute VB_Name = "InImpl"
Option Explicit

' Fast-path membership checks for known COM container types by TypeName.
' Sets matched=True and returns True/False if type is known with O(1) membership.
' Sets matched=False for known types without O(1) membership (Collection, ListBox, etc.)
' or for unknown types (caller falls back to iteration or dynamic check).
Public Function PyInByTypeName(ByRef item As Variant, ByRef container As Object, _
                                ByRef matched As Boolean) As Boolean
    matched = True
    Select Case TypeName(container)
        ' --- Scripting.Dictionary — .Exists(key) ---
        Case "Dictionary"
            PyInByTypeName = container.Exists(item)
        
        ' --- .NET System.Collections ---
        ' Hashtable uses .ContainsKey  (IDictionary)
        Case "Hashtable"
            PyInByTypeName = container.ContainsKey(item)
        
        ' SortedList implements IDictionary — .ContainsKey
        Case "SortedList"
            PyInByTypeName = container.ContainsKey(item)
        
        ' ICollection / IList types — .Contains(item)
        Case "ArrayList", "ICollection", "BitArray", "Stack", "Queue"
            PyInByTypeName = container.Contains(item)
        
        ' --- Known types without O(1) membership — signal matched=False ---
        Case "Collection"
            matched = False
        Case "ListBox", "ComboBox", "DirListBox", "FileListBox"
            matched = False
        Case "Form", "MDIForm", "PropertyPage"
            matched = False
        Case "Recordset"
            matched = False
        
        ' --- Unknown type — signal caller to try dynamic or iteration ---
        Case Else
            matched = False
    End Select
End Function

' Tries late-bound Contains, ContainsKey, and Exists on unknown object types.
' Returns Variant tri-state:
'   True / False  — authoritative answer (container supports one of these methods)
'   Empty         — none of the methods exist, caller should use iteration
Public Function PyInDynamicCheck(ByRef item As Variant, ByRef container As Object) As Variant
    On Error Resume Next
    
    PyInDynamicCheck = container.Contains(item)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    
    PyInDynamicCheck = container.ContainsKey(item)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    
    PyInDynamicCheck = container.Exists(item)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    
    PyInDynamicCheck = Empty
End Function
