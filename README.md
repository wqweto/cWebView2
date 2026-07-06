# cWebView2 — RC6-compatible WebView2 wrapper for VB6

A drop-in replacement of vbRichClient6's `cWebView2` class, implemented in pure VB6
directly against the official Microsoft WebView2 COM API. No `RC6.dll` dependency,
no dependency on the WebView2 SDK typelib — the only redistributable is
`WebView2Loader.dll` deployed next to the exe, targeting the Evergreen WebView2
Runtime installed on the machine.

## Layout

| Path | Purpose |
|---|---|
| `src/cWebView2.cls` | The public class — RC6-compatible properties, methods and events |
| `src/cWebView2Callback.cls` | Implements every native callback interface (completed-handlers, event handlers, environment options) and forwards to `cWebView2`'s Friend sinks |
| `typelib/VBWebView2Impl.odl` / `.tlb` | Hand-written companion type library with VB6-safe redeclarations of the WebView2 interfaces (see below) |
| `test/WebView2Loader.dll` | Microsoft's loader shim (x86), resolves the installed Evergreen runtime |
| `test/Form1.frm` | Manual test harness, grown alongside the implementation |
| `test/Project1.vbp` | Test exe project — references only `stdole2.tlb` and `VBWebView2Impl.tlb` |
| `doc/WebView2.idl` | The official WebView2 SDK IDL, kept as read-only reference only |
| `doc/RC6.idl` | vbRichClient6 IDL, kept as API-surface reference only |

## How it works

VB6 cannot `Implements` (or in places even call) the WebView2 SDK interfaces
because they use automation-hostile types: raw `BOOL`, `LPWSTR`/`LPCWSTR`,
unsigned ints, and even plain `IUnknown*` parameters. `VBWebView2Impl.odl`
redeclares every interface the code touches under an `IVB...` name with the
**exact same IID** as the SDK original — so `QueryInterface` and vtable calls
to/from native code resolve correctly — but with VB6-safe types:

- `BOOL` → `long` (0 = False, nonzero = True)
- `LPWSTR`/`LPCWSTR` → `long` (raw pointer, converted with `StrPtr` /
  `SysReAllocString` helpers; strings the API documents as caller-owned are
  freed with `CoTaskMemFree`)
- `UINT32`/`UINT64`/enums/`HRESULT` params → `long`
- interface pointers not otherwise needed → `VB6IUnknown`, a locally-declared
  marker interface binary-identical to `IUnknown` (VB6 rejects the imported
  stdole `IUnknown` as a parameter type in `Implements` members)
- interface version chains (`ICoreWebView2_2..11`, `ICoreWebView2Settings2..6`,
  `ICoreWebView2Controller` unchanged) are flattened onto their base vtable
  and published under the newest IID actually needed
- unused members that precede needed ones are still declared, purely to keep
  vtable slots aligned
- every interface is forward-declared once near the top of the file (`mktyplib`
  supports bare `interface IFoo;` declarations), so any interface can reference
  any other regardless of which one is fully defined first further down

The tlb is compiled with `MKTYPLIB.EXE VBWebView2Impl.odl` (VC98).

RC6's synchronous call style (`BindTo`, `Navigate`, `jsRun`, ... block until done
or timeout) is reproduced by spinning the thread message pump
(`PeekMessage`/`DispatchMessage` + `CoWaitForMultipleHandles`) until the native
async completion fires. Completions are correlated by `Currency` tokens so
nested and interleaved calls cannot cross wires.

## RC6 compatibility notes

- The instance registers its `cWebView2Callback` as host object `vbHost` and
  injects a global `vbH()` accessor on document creation, so existing
  RC6-style page script like `vbH().RaiseMessageEvent('title_change',
  document.title)` works unchanged. The callback only forwards a handful of
  methods (`RaiseMessageEvent`, `RaiseContextMenuEvent`) to the owning
  `cWebView2` — script never obtains a direct reference to the `cWebView2`
  instance itself. A `contextmenu` listener is injected the same way to
  raise `UserContextMenu`.
- Default user data folder is `%LOCALAPPDATA%\<EXENAME>`, matching RC6, instead
  of the native default `<exe>.WebView2` next to the executable.
- `JSMessage` fires from both the RC6 channel (`vbH().RaiseMessageEvent`) and
  native `chrome.webview.postMessage`; JSON object content is decoded into the
  `oJSONContent` parameter.
- `AddObject "Name", Obj` also publishes a global `window.Name` alias of the
  synchronous host-object proxy, so page script can call `Name.MethodName(...)`
  directly, like with RC6. The native `chrome.webview.hostObjects[.sync].Name`
  forms remain available (the async proxy returns Promises).

### Deliberate deviations from RC6

- `CapturePreview` returns a `Byte()` array with the raw PNG/JPEG data (RC6
  returned a Cairo surface).
- `WebResourceRequested` is a simplified allow/deny gate — deny answers the
  request with a synthesized `403`, no response-body rewriting.
- `GotFocus`/`LostFocus` `Reason` is a best-effort heuristic (primed by
  `SetFocus`, defaults to `FocusReason_PROGRAMMATIC`).
- `CallDevToolsProtocolMethod` blocks and returns the CDP response, decoded
  like `jsRun`'s result, instead of RC6's fire-and-forget signature (RC6's own
  IDL has no retval either, but discarding the CDP response is rarely useful).
- `BindTo`'s `AllowSingleSignOnUsingOSPrimaryAccount` is `Boolean`, not RC6's
  `Long`.
- `BindTo` has three trailing optional parameters beyond RC6's signature —
  `ExclusiveUserDataFolderAccess`, `IsCustomCrashReportingEnabled`,
  `EnableTrackingPrevention` (default `True`, matching WebView2's own default)
  — surfacing `ICoreWebView2EnvironmentOptions2/3/5`. `Options4`'s custom
  scheme registration is not exposed (no RC6 counterpart, different shape of
  feature entirely — an array of registration objects, not a simple flag).

## Building

1. `MKTYPLIB.EXE typelib\VBWebView2Impl.odl` whenever the odl changes.
2. Build `test\Project1.vbp` with VB6 (`VB6.EXE /make`).
3. Ship `WebView2Loader.dll` next to the compiled exe. The Evergreen WebView2
   Runtime must be installed on the target machine (preinstalled on Windows 11).

## Status

All RC6 `cWebView2` properties, methods and events are implemented and verified
against live pages (navigation, JS interop incl. blocking/async calls and JSON
marshaling, host objects, web messages, settings, script dialogs, permissions,
new-window, accelerator keys, focus, web-resource filtering, response
introspection, frames, downloads, capture) — except `GetMostRecentInstallPath`
and `jsCallByName`, which intentionally raise "not yet implemented" (Evergreen
deployment resolves the runtime automatically instead of needing an install
path; WebView2 cannot proxy live JS objects the way RC6's `Obj` parameter
implies).

### Additions beyond RC6 parity

A handful of native WebView2/`WebBrowser`-style events with no RC6 counterpart
are also exposed, since the native plumbing is either free (already needed for
something else) or fills a gap the JS-bridge-based events can't:

- `ContentLoading(IsErrorPage)` — fires before `NavigationCompleted`, akin to
  `WebBrowser`'s early loading moment.
- `HistoryChanged()` — back/forward stack changed; pair with `CanGoBack`/`CanGoForward`.
- `WindowCloseRequested()` — page called `window.close()`.
- `ZoomFactorChanged()` — native counterpart to the `ZoomFactor` property.
- `MoveFocusRequested(Reason, Handled)` — Tab/Shift+Tab reached the edge of the
  web content; `Reason` reuses `eWebView2FocusReason`.
- `ContextMenuRequested(PageURI, LinkURI, SelectionText, ScreenX, ScreenY, Handled)`
  — the *native* context-menu event, richer than the JS-injected `UserContextMenu`
  (which only reports coordinates): carries link/selection info and can actually
  suppress the browser's context menu via `Handled`. Deliberately simplified —
  no custom menu item injection/`CustomItemSelected`, matching the
  `WebResourceRequested` simplified-gate precedent. `UserContextMenu` is
  unchanged and still fires alongside it.

`ContextMenuRequested` required flattening `ICoreWebView2_5`..`_11` (only
`_11`'s `ContextMenuRequested` is used; the rest are unused placeholders kept
solely to preserve vtable alignment), same technique as the earlier `_4` flatten.
