# AGENTS.md - BASICpy Codebase Guidelines

This file provides guidelines for agentic coding assistants working in this Visual Basic 6 (VB6) codebase.

## Project Overview

This is a VB6 ActiveX DLL project (`BASICpy.vbp`) that ports Python's standard library builtin functions to VB Classic. It follows Python semantics closely while accommodating VB6 constraints.

### Architecture

- **Builtin.cls**: Global namespace class (`VB_GlobalNameSpace = True`, `VB_PredeclaredId = True`) acting as public facade — thin (2-4 line) Public Functions that delegate to implementation modules
- **Builtin/ directory**: All implementation modules (.bas) and class files (.cls)
  - **.bas modules**: Implementation helpers (LenImpl, IntImpl, InImpl, IterImpl, AbsImpl, AllAnyImpl, RangeImpl, NextImpl, BinOctHexImpl, PyBoolImpl, VBArrayCoreImpl, IntCoreImpl, InCoreImpl, IterCoreImpl, TupleImpl)
  - **Python protocol interfaces**: IBoolean, ILength, IIndex, IAbsolute, IPredicate, IsTruthy, ISupplier, ISequence, IIterable, IIterator, IContains, IMutableSequence
  - **Implementation classes**: ArrayIterator, ArrayRevIterator, SequenceIterator, StringIterator, SupplierIterator, RangeIterator, Tuple, VBArray, EnumVARIANTIterator, IEnumVARIANTWrapper
- **General utilities**: Ensure.cls (validation), PyMath.cls (math), CBasicPyTypes.cls, time.cls, ErrorCodes.bas
- **PInvoke/**: Win32 API declarations (DispCall.bas — DispCallFunc from oleaut32.dll)

## Build Commands

**Windows-only build environment required** — This is a legacy VB6 project that requires the Visual Basic 6.0 IDE for compilation.

**CRITICAL RULE**: Do not build or test this project automatically without explicit instructions from the user.

## Scripts

Located in `scripts/` directory:

- **`fix-line-endings.sh <project-root>`** — Converts line endings of `.cls`, `.bas`, and `.vbp` files to CRLF (Windows/DOS) using `unix2dos`. Run after editing files to ensure VB6 compatibility.
  ```bash
  bash scripts/fix-line-endings.sh .
  ```

- **`check_decl_style.py`** — Scans VB6 source files for declaration style violations: `Dim X As Type: X = value` patterns where `X` is not used within 5 lines (indicating the declaration+initialization is too far from first use). Also checks for `Dim` statements inside control blocks (`If`/`Else`/`Select Case`/`For`/`While`).
  ```bash
  python3 scripts/check_decl_style.py
  ```
  Output lists each violating line with file path, line number, and variable name. Zero output means no violations.

## Code Style Guidelines

### Naming Conventions
- **Classes/Modules**: PascalCase (e.g., `ArrayIterator`, `LenImpl`)
- **Methods/Functions**: PascalCase (e.g., `PyBool`, `PyLenByTypeOf`)
- **Variables**: PascalCase (e.g., `m_start`, `m_length`)
- **Private Fields**: Prefix with `m_` (e.g., `m_data`, `m_index`)
- **Constants**: ALL_CAPS (e.g., `StopIterationCode.StopIteration`)

### Formatting
- **Indentation**: 4 spaces
- **Line Length**: Aim for 80-100 characters maximum
- **Variable Declarations**: Use `Option Explicit` at the top of every file
- **Declare/Assign at Point of Use**: Place each `Dim` at procedure scope just before its first use — not grouped at the top
  - VB6 has **no block scope**: all `Dim` statements are hoisted to procedure scope regardless of where they appear
  - Never put `Dim` inside `If`/`Else`/`Select Case`/`For`/`While` blocks — it creates a misleading illusion of block scoping and causes duplicate-declaration errors
  - If a variable is used across multiple branches, declare once at procedure scope and assign with `Set`/`=` in each branch
- **Inline Declaration + Assignment**: Combine `Dim` with its first assignment on one line only when tightly coupled AND the variable is used immediately after (e.g., `Dim lenImpl As ILength: Set lenImpl = value`). If a variable is declared but assigned later (conditionally or in another branch), keep the `Dim` separate and place it close to its first actual assignment. Never declare+initialize a variable far before its first use — move both declaration and initialization to where the variable is actually read or passed
- **Object vs value assignment**: Use `Set` for object references (e.g., `Set x = obj`), `=` for value types (e.g., `x = 42`)

### Language Features
- **VB6 Compatibility**: Target Visual Basic 6.0 runtime
- **Data Types**: Use appropriate VB6 types (Long, Integer, String, Variant)
- **Error Handling**: Use `On Error` statements and `Ensure.cls` validation

### ByRef/ByVal Convention
- **ByRef** on Variant parameters that can receive arrays (avoids SafeArrayCopy crash on NULL SAFEARRAY)
- **ByVal** on non-array parameters (scalars, strings, objects)
- See `PyBool` (ByRef), `PyIter(value)` (ByRef), `PyRange(start/stop/step)` (ByVal), `PyInt(value)` (ByVal)

### Variant Object Reference Safety (IsObject/Set Dispatch)
When assigning from any Variant expression that may hold an object reference (whether from an array element, function return, collection item, or ByRef parameter), you **must** dispatch between `Set` and `=` based on `IsObject()`:

```vb
' Use a single-line If to avoid End If:
If IsObject(source) Then Set dest = source _
                                               Else dest = source
```

**Why**: VB6's `Let` assignment (`=`) between Variants does not call `AddRef` on the underlying COM object when the source Variant contains an object reference. Without `Set`, the object can be destroyed prematurely when the source reference goes out of scope, leaving a dangling pointer.

- `IsObject` on a value-type Variant (`vbLong`, `vbString`, etc.) returns `False` → uses `=` (Let), correct for values
- `IsObject` on an object-type Variant (`vbObject`, `vbDispatch`) returns `True` → uses `Set`, which calls `AddRef` (see `ArrayIterator.cls:45-46`)
- `IsObject` on `Nothing` returns `True` → uses `Set`, which correctly propagates `Nothing`
- This applies to **all** Variant-to-Variant assignments: array element reads, function return values (`FunctionName = expr`), ByRef output parameters (`param = expr`), Collection items, and intermediary local variables
- **Prefer `LetSet` helper**: Instead of writing the inline `If IsObject...` pattern, call the `LetSet` procedure (inside `Builtin.cls`) or `Builtin.LetSet` (from outside the class) for cleaner code. The `LetSet` procedure implements the same IsObject/Set dispatch logic.

### Code Quality Guidelines

#### Error Handling
```vb
' Use custom error handling with Ensure.cls
Private Sub SomeMethod()
    On Error GoTo ErrorHandler
    ' Code here
    Ensure.IsTrue condition, ErrorCodes.TypeMismatch, "SomeMethod", "Error message"
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
```

#### Comments
- Use single quote comments for explanations
- Document Python analog where behavior matches: `' Matches Python's bool(value)`
- Document VB6-specific accommodations: `' ByRef to avoid SafeArrayCopy on NULL SAFEARRAY`
- Use numbered steps for dispatch logic (0-based)
- Use `' ─── Section headers ───` for section breaks
- Use `' TODO:` for known gaps

## Agent Instructions

When working in this codebase, agents should:

1. **Match Python semantics** — follow Python's behavior for builtin functions, document divergences
2. **Follow existing code style** — 4-space indent, PascalCase, `m_` prefix
3. **Use Ensure.cls** for parameter validation with descriptive messages
4. **Use ByRef** for array-accepting Variant parameters to avoid SafeArrayCopy issues
5. **Handle Nothing** on all entry-point functions (PyBool returns False, PyLen/PyIter raise)
6. **Use the object-safe variant read pattern** (`IsObject` check with `Set`/`=` dispatch)
7. **Keep Builtin.cls as thin facades** — implementation goes in `.bas` modules under Builtin/
8. **Never attempt automated builds or tests** without explicit user instruction