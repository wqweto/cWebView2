# RC6 cWebView2 replacement

## Coding Style for this Module

- All private helpers are `Private Function` / `Private Sub`
- Use `On Error GoTo EH` when necessary to Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
- Use `Long` for all pixel values and array indices
- Avoid `Variant`; avoid dynamic arrays in inner loops (pre-allocate)
- One blank line between procedures
- One blank line between the `Dim` block and the first executable line inside a procedure; no other blank lines inside a procedure body
- Comments only where the VB6 diverges non-obviously from the C++ (e.g. unsigned workarounds)
- Use hungarian notation: s - String, l - Long, n - Integer, b - Boolean, o - Object, c - Collection, d - Date, dbl - Double, sng - Single, byt - Byte, u - UDTs, h - Handles (incl. hResult), cy - Currency
- Use `ba` prefix for byte arrays, use `a` only for arrays no matter what type they are
- Use `m_` prefix for member variables and `g_` for global ones
- Use `md` prefix for standard modules, `c` for classes, `frm` for forms, `ctx` for user-controls
- Declare all variable at the beginning of procedure and separate with a blank line from code
- Single local variable declaration per line, align variable's data-type at column 25
- Align API consts data-types at column 45
- Align module variables data-types at column 37
- Declare API consts local to a routine if not used in any other routine
- API declares use "dllname" without .dll suffix, always Unicode versions (aliased to names without W)
- Don't put any blank lines in code, use comments instead for separators
- Use '--- for comments start instead of single ' unless a comment banner at start of procedure/module
- Put only one statement per logical line i.e. don't use : to separate multiple statements
- Put `If` statements on separate lines i.e. don't merge `If Cond Then Stmt` on a single line
- Use explicit `Call` only with API functions which have result discarded i.e. not used in If statements; don't use `Call` statement otherwise
- Use `QH` label for cleanup before `EH` label
- Align `Case`es after `Select Case` at the same column i.e. don't indent
- All `Long` const hex literals between &HA000 and &HFFFF must use & type character i.e. &HA000& or risk being sign extended
- Never use `Next Var` i.e. just `Next`
- Always omit `ByRef` parameter passing type, only specify `ByVal` if needed
- Order of procedures in module: public events/enums/types, API declares, member variables and private enums/types, properties, methods, event handlers, base class events (e.g. Class_Terminate)
- Order by type within properties/methods: public, friend and private
- Use `pv` prefix for private procedures and `fr` for friend ones
