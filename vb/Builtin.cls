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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ─── Public API facades (implementations in Builtin/ modules) ───
' All functions mirror Python builtins with VB6-specific accommodations
' (ByRef for array safety, Optional for IsMissing detection, etc.).
' Design rules:
'   ByRef = array-capable (avoids SafeArrayCopy crash on undimensioned arrays)
'   ByVal = non-array (numbers, strings, objects only)

' Python: len(value). VB6: ByRef avoids SafeArrayCopy on NULL SAFEARRAY
' when value is a typed dynamic array passed across COM boundary.
' The >= 0 return guarantees success — on failure the Core raises.
Public Function PyLen(ByRef value As Variant) As Long
Attribute PyLen.VB_Description = "Returns the length of a value matching Python's len().\r\nVB6: Negative return (-1) indicates the value has no len().\r\nArgs:\r\n  value (Variant): Array, string, or object implementing ILength/Count/Length.\r\nReturns:\r\n  Long: Length or -1 on failure (caller raises)."
    PyLen = PyLenCore(value)
End Function

' Python: abs(value). VB6: ByVal because PyAbs never accepts arrays
' (only numeric scalars or IAbsolute objects).
Public Function PyAbs(ByVal value As Variant) As Variant
Attribute PyAbs.VB_Description = "Returns the absolute value of a number, matching Python's abs().\r\nVB6: Accepts numeric scalars or objects implementing IAbsolute. Arrays raise.\r\nArgs:\r\n  value (Variant): Numeric value or IAbsolute object.\r\nReturns:\r\n  Variant: The absolute value."
    PyAbs = PyAbsCore(value)
End Function

' Python: all(iterable). VB6: ByRef for array safety. Optional IPredicate
' allows custom truth testing; defaults to IsTruthy (Python's default bool()).
' Equivalent: all(x > 0 for x in seq) via predicate.
Public Function PyAll(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
Attribute PyAll.VB_Description = "Returns True if all elements of the iterable are truthy (or pass the predicate), matching Python's all().\r\nArgs:\r\n  value (Variant): Iterable (array, string, IIterable, COM collection).\r\n  predicate (IPredicate): Optional custom truth test. Defaults to IsTruthy.\r\nReturns:\r\n  Boolean: True if all elements pass."
    PyAll = PyAllCore(value, predicate)
End Function

' Python: any(iterable). VB6: ByRef for array safety. Same predicate
' semantics as PyAll. Short-circuits on first True.
Public Function PyAny(ByRef value As Variant, Optional ByVal predicate As IPredicate) As Boolean
Attribute PyAny.VB_Description = "Returns True if any element of the iterable is truthy (or passes the predicate), matching Python's any().\r\nShort-circuits on first True.\r\nArgs:\r\n  value (Variant): Iterable (array, string, IIterable, COM collection).\r\n  predicate (IPredicate): Optional custom truth test. Defaults to IsTruthy.\r\nReturns:\r\n  Boolean: True if any element passes."
    PyAny = PyAnyCore(value, predicate)
End Function

' Python: range([start,] stop[, step]). VB6: ByVal args because PyRange
' never accepts arrays (integer scalars only). Returns RangeIterator
' (ISequence+ILength+IIterable) directly — O(1) length, lazy iteration.
' Default step=1 matches Python. step arg is ByVal Long (non-Optional
' with default) so always present.
Public Function PyRange(Optional ByVal arg1 As Variant, _
                        Optional ByVal arg2 As Variant, _
                        Optional ByVal arg3 As Long = 1) As RangeIterator
Attribute PyRange.VB_Description = "Creates a lazy RangeIterator matching Python's range().\r\nArgs:\r\n  arg1 (Variant): Stop (1-arg) or start (2-arg).\r\n  arg2 (Variant): Stop (2-arg).\r\n  arg3 (Long): Step (default 1).\r\nReturns:\r\n  RangeIterator: ISequence+ILength+IIterable with O(1) length."
    Set PyRange = PyRangeCore(arg1, arg2, arg3)
End Function

' Python: next(iterator[, default]). VB6: ByRef IIterator (not Variant)
' because next() only works on iterators, not generic iterables.
' Exhaustion detected via sentinel value (IsError), matching Python's
' StopIteration convention. One-arg form raises; two-arg returns default.
Public Function PyNext(ByRef iter As IIterator, Optional ByVal default As Variant) As Variant
Attribute PyNext.VB_Description = "Returns the next item from an iterator, matching Python's next(it[, default]).\r\nVB6: Exhaustion detected via sentinel value (IsError).\r\nArgs:\r\n  iter (IIterator): The iterator to advance.\r\n  default (Variant): Optional default if iterator exhausted.\r\nReturns:\r\n  Variant: Next item or default. Raises if exhausted with no default."
    PyNext = PyNextCore(iter, default)
End Function

' Python: safe iterator loop helper (no direct Python equivalent but
' idiomatically replaces the prime-then-loop pattern).
' VB6: ByRef item receives the value, ByRef iter avoids interface copy.
' Returns True if item valid, False if exhausted.
' Usage: Dim item: Do While PyTryNext(iter, item)
Public Function PyTryNext(ByRef iter As IIterator, ByRef item As Variant) As Boolean
Attribute PyTryNext.VB_Description = "Safe iterator loop helper. Calls NextItem and returns True if valid.\r\nVB6: Replaces the prime-then-loop pattern.\r\nArgs:\r\n  iter (IIterator): The iterator to advance.\r\n  item (Variant): Receives the next value.\r\nReturns:\r\n  Boolean: True if item valid, False if exhausted.\r\nUsage:\r\n  Dim item: Do While PyTryNext(iter, item): ... Loop"
    PyTryNext = PyTryNextCore(iter, item)
End Function

' Python: tuple([iterable]). VB6: ByRef for array safety. Returns Tuple
' as Variant (object reference). Existing Tuple is returned unchanged
' (identity preservation). Strings are iterated character by character.
Public Function PyTuple(Optional ByRef value As Variant) As Variant
Attribute PyTuple.VB_Description = "Creates an immutable Tuple from a value, matching Python's tuple().\r\nVB6: Existing Tuple is returned unchanged (identity preservation).\r\nArgs:\r\n  value (Variant): Array, string, iterable, or omitted for empty ().\r\nReturns:\r\n  Variant: Tuple object reference."
    PyTuple = PyTupleCore(value)
End Function

' Python: bin(value). VB6: ByVal (integer only, no array). Accepts Byte,
' Integer, Long, or IIndex objects. Output format: "-0bNNN" for negative,
' "0bNNN" for positive. Special-cases Long.MinValue (&H80000000).
Public Function PyBin(ByVal value As Variant) As String
Attribute PyBin.VB_Description = "Converts an integer to binary with '0b' prefix, matching Python's bin().\r\nVB6: Accepts Byte, Integer, Long, or IIndex objects. Output: '-0bNNN' for negative.\r\nArgs:\r\n  value (Variant): Integer value or IIndex object.\r\nReturns:\r\n  String: Binary representation with '0b' prefix."
    PyBin = PyBinCore(value)
End Function

' Python: oct(value). VB6: ByVal (integer only, no array). Output format:
' "-0oNNN" / "0oNNN". Same type restrictions as PyBin.
Public Function PyOct(ByVal value As Variant) As String
Attribute PyOct.VB_Description = "Converts an integer to octal with '0o' prefix, matching Python's oct().\r\nVB6: Accepts Byte, Integer, Long, or IIndex objects. Output: '-0oNNN' for negative.\r\nArgs:\r\n  value (Variant): Integer value or IIndex object.\r\nReturns:\r\n  String: Octal representation with '0o' prefix."
    PyOct = PyOctCore(value)
End Function

' Python: hex(value). VB6: ByVal (integer only, no array). Output format:
' "-0xNNNN" / "0xNNNN" (lowercase hex, matching Python). Same type
' restrictions as PyBin.
Public Function PyHex(ByVal value As Variant) As String
Attribute PyHex.VB_Description = "Converts an integer to lowercase hex with '0x' prefix, matching Python's hex().\r\nVB6: Accepts Byte, Integer, Long, or IIndex objects. Output: '-0xNNNN' for negative.\r\nArgs:\r\n  value (Variant): Integer value or IIndex object.\r\nReturns:\r\n  String: Hex representation with '0x' prefix."
    PyHex = PyHexCore(value)
End Function

' Python: bool([value]). VB6: Optional ByRef allows no-arg call → False.
' IsMissing is guarded behind IsArray check because IsMissing throws on
' undimensioned arrays. Dispatch order: Nothing→False, missing/empty/null→False,
' Boolean→self, numeric→nonzero, length→(len>0), IBoolean→.__bool__(),
' ILength→len>0, objects→True, fallback→CBool or True.
Public Function PyBool(Optional ByRef value As Variant) As Boolean
Attribute PyBool.VB_Description = "Returns the truth value matching Python's bool().\r\nVB6: Dispatch order — Nothing→F, missing/empty/null→F, Boolean→self, numeric→nonzero, length→len>0, IBoolean.__bool__(), ILength.__len__()>0, objects→T, fallback→CBool or T.\r\nArgs:\r\n  value (Variant): Optional value or omitted (False).\r\nReturns:\r\n  Boolean: Truth value."
    PyBool = PyBoolCore(value)
End Function

' Python: list([iterable]) in VB6 terms — wraps values as a mutable,
' 0-based Variant array (VBArray). VB6: ByRef for array safety.
' Existing VBArray returned via Copy(). Strings iterated char by char.
' Iterables collected via iteration.
Public Function PyVBArray(Optional ByRef value As Variant) As VBArray
Attribute PyVBArray.VB_Description = "Creates a mutable 0-based Variant array (VBArray), akin to Python's list().\r\nVB6: Existing VBArray is copied. Strings iterated character by character.\r\nArgs:\r\n  value (Variant): Array, string, iterable, or omitted for empty.\r\nReturns:\r\n  VBArray: Mutable sequence object."
    Set PyVBArray = PyVBArrayCore(value)
End Function

' Python: int([x[, base]]). VB6: ByVal because arrays raise TypeMismatch
' (int() never accepts arrays, so no SafeArrayCopy risk). IsArray guard
' placed before IsMissing to avoid IsMissing throw on undimensioned arrays.
' Float truncation via VB6's Fix() matches Python's int(1.9) → 1.
' String parsing with prefix detection (0x/0o/0b), bases 2-36.
Public Function PyInt(Optional ByVal value As Variant, Optional ByVal base As Variant) As Variant
Attribute PyInt.VB_Description = "Converts a value to an integer, matching Python's int().\r\nVB6: Float truncation via Fix(). String parsing with 0x/0o/0b prefix detection, bases 2-36. Arrays raise TypeMismatch.\r\nArgs:\r\n  value (Variant): Number, string, IIndex object, or omitted (0).\r\n  base (Variant): Optional base (2-36) for string parsing.\r\nReturns:\r\n  Variant: Integer value."
    PyInt = PyIntCore(value, base)
End Function

' Python: x in container. VB6: Both item and container ByRef (container
' may be an array). Dispatch: IContains protocol → string substring (InStr)
' → TypeName-matched fast paths (Dictionary.Exists, ArrayList.Contains, etc.)
' → iteration fallback → dynamic IDispatch trial (last resort).
' VB6 arrays: undimensioned = empty (nothing to find).
Public Function PyIn(ByRef item As Variant, ByRef container As Variant) As Boolean
Attribute PyIn.VB_Description = "Tests membership, matching Python's 'x in container'.\r\nVB6: Both item and container ByRef for array safety. Dispatch: IContains → InStr substring → TypeName fast paths → iteration → dynamic IDispatch.\r\nArgs:\r\n  item (Variant): The value to search for.\r\n  container (Variant): The container (array, string, object).\r\nReturns:\r\n  Boolean: True if item is found."
    PyIn = PyInCore(item, container)
End Function

' Python: iter(iterable[, sentinel]). VB6: ByRef for array safety (arrays
' are valid iterables). Single-arg: tries IIterable, then ArrayIterator,
' StringIterator, SequenceIterator, IEnumVARIANT. Two-arg: ISupplier +
' sentinel (calls supplier until it returns sentinel). Nothing raises
' "'Nothing' object is not iterable".
Public Function PyIter(ByRef value As Variant, Optional ByVal sentinel As Variant) As IIterator
Attribute PyIter.VB_Description = "Returns an iterator for the given object, matching Python's iter().\r\nVB6: Single-arg: IIterable → ArrayIterator → StringIterator → SequenceIterator → IEnumVARIANT. Two-arg: ISupplier + sentinel. Nothing raises.\r\nArgs:\r\n  value (Variant): Iterable (array, string, IIterable, COM collection).\r\n  sentinel (Variant): Optional sentinel for callable-based iteration.\r\nReturns:\r\n  IIterator: Iterator object."
    Set PyIter = PyIterCore(value, sentinel)
End Function
