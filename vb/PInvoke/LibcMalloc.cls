VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibcMalloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' All aligned-* exports have "_" as part of the name, so mangle=False.

' void _aligned_free(void* memblock)
Public Sub aligned_free(ByVal memblock As Long)
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_free", False)
    Dim args(0 To 0) As Variant
    args(0) = memblock
    Dim result As Variant
    fn.CallFunc vbEmpty, result, args
End Sub

' void* _aligned_malloc(size_t size, size_t alignment)
Public Function aligned_malloc(ByVal size As Long, ByVal alignment As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_malloc", False)
    Dim args(0 To 1) As Variant
    args(0) = size
    args(1) = alignment
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    aligned_malloc = result
End Function

' size_t _aligned_msize(void* memblock, size_t alignment, size_t offset)
Public Function aligned_msize(ByVal memblock As Long, _
                              ByVal alignment As Long, _
                              ByVal offset As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_msize", False)
    Dim args(0 To 2) As Variant
    args(0) = memblock
    args(1) = alignment
    args(2) = offset
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    aligned_msize = result
End Function

' void* _aligned_offset_malloc(size_t size, size_t alignment, size_t offset)
Public Function aligned_offset_malloc(ByVal size As Long, ByVal alignment As Long, _
                                      ByVal offset As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_offset_malloc", False)
    Dim args(0 To 2) As Variant
    args(0) = size
    args(1) = alignment
    args(2) = offset
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    aligned_offset_malloc = result
End Function

' void* _aligned_offset_realloc(void* memblock, size_t size, size_t alignment, size_t offset)
Public Function aligned_offset_realloc(ByVal memblock As Long, ByVal size As Long, _
                                       ByVal alignment As Long, ByVal offset As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_offset_realloc", False)
    Dim args(0 To 3) As Variant
    args(0) = memblock
    args(1) = size
    args(2) = alignment
    args(3) = offset
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    aligned_offset_realloc = result
End Function

' void* _aligned_offset_recalloc(void* memblock, size_t num, size_t size, size_t alignment, size_t offset)
Public Function aligned_offset_recalloc(ByVal memblock As Long, ByVal num As Long, _
                                        ByVal size As Long, ByVal alignment As Long, _
                                        ByVal offset As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_offset_recalloc", False)
    Dim args(0 To 4) As Variant
    args(0) = memblock
    args(1) = num
    args(2) = size
    args(3) = alignment
    args(4) = offset
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    aligned_offset_recalloc = result
End Function

' void* _aligned_realloc(void* memblock, size_t size, size_t alignment)
Public Function aligned_realloc(ByVal memblock As Long, ByVal size As Long, _
                                ByVal alignment As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_realloc", False)
    Dim args(0 To 2) As Variant
    args(0) = memblock
    args(1) = size
    args(2) = alignment
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    aligned_realloc = result
End Function

' void* _aligned_recalloc(void* memblock, size_t num, size_t size, size_t alignment)
Public Function aligned_recalloc(ByVal memblock As Long, ByVal num As Long, _
                                 ByVal size As Long, ByVal alignment As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_aligned_recalloc", False)
    Dim args(0 To 3) As Variant
    args(0) = memblock
    args(1) = num
    args(2) = size
    args(3) = alignment
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    aligned_recalloc = result
End Function

' WARNING: _alloca allocates on the stack of the calling function.
' Through DispCallFunc, it allocates on DispCallFunc's stack frame, so the
' returned pointer is INVALID after this function returns. Use only if you
' understand this limitation (e.g., for testing the call mechanism).
' void* _alloca(size_t size)
Public Function alloca(ByVal size As Long) As Long
    Dim fn As CDeclLibraryFunction: Set fn = LibcLibrary.ResolveFunction("_alloca", False)
    Dim args(0 To 0) As Variant
    args(0) = size
    Dim result As Variant
    fn.CallFunc vbLong, result, args
    alloca = result
End Function
