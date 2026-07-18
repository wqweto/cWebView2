# RC6.dll string dump around SetFuncObj (the synchronous JS bridge)

Extracted from `C:\Work\ExternalComponents\RC6\RC6.dll` (4,476,416 bytes,
2023-07-22) with `strings -e l -t d` (UTF-16LE literals, decimal byte
offsets) and `strings -t d` (ANSI). This is the compiled-in glue behind
RC6's `jsRun`/`jsProp`/`jsCallByName` - the mechanism our `pvBridgeInvoke`
replicates.

## 1. The injected JS bridge (reassembled in offset order)

String literals at offsets 747440-751224, in the order they appear. The
line fragments interleave (VB6 string constants are pooled), reassembled
by function:

```js
function vbH(){try{return chrome.webview.hostObjects.sync.vb_Host} catch(e){}}

function RunFunc(funcName, token) {
  var ctx  = window;
  var ns   = funcName.split('.'), func = ns.pop();
  var args = Array.prototype.slice.call(arguments, 2);
  for(var i= 0; i < ns.length; i++){ ctx = ctx[ns[i]] }
  var eMsg = '', res = null;
  try{res  = ctx[func].apply(ctx, args)} catch(e){eMsg='Err in '+func+'(): '+e.message};
  try{vbH().RaiseResultEvent(res, token, eMsg);} catch(e){}
}

function Cbn(obj, callType, methodName, token) {
  try{switch (callType) {
    case 1: res = obj[methodName].apply(obj, Array.prototype.slice.call(arguments, 4));break;
    case 2: try{vbH().RaiseMessageEvent (methodName, JSON.stringify(obj)); res = obj[methodName]} catch(e){};break;
    case 4: res = obj[methodName]=arguments[4];break;
  }} catch(e){eMsg='Err in '+methodName+'(): '+e.message};
  // (RaiseResultEvent tail shared with the other functions)
}

function GetProp(propName, token){
  var ctx = window, ns   = propName.split('.'), prop = ns.pop();
  try{res = ctx[prop]} catch(e){eMsg='Err in '+propName+'-Get: '+e.message};
  // (RaiseResultEvent tail)
}

function LetProp(propName, RHS, token){
  try{ctx[prop]=RHS} catch(e){eMsg=e.message};
  try{vbH().RaiseResultEvent(true, token, eMsg);} catch(e){}
}

document.addEventListener('contextmenu', function(e){
  try{vbH().RaiseMessageEvent('contextmenu', JSON.stringify({x:e.screenX, y:e.screenY})) }catch(e){}
})  // closing fragment pooled elsewhere

window.onblur =  function(){ try{vbH().RaiseMessageEvent('onblur', '')} catch(e){} }
window.onfocus = function(){ try{vbH().RaiseMessageEvent('onfocus','')} catch(e){} }

try{cbFuncs={run:RunFunc, cbn:Cbn, propGet:GetProp, propLet:LetProp}; vbH().SetFuncObj(cbFuncs);} catch(e){}
```

Raw offsets:

```
747440 vb_Host
747460   var ns   = funcName.split('.'), func = ns.pop();
747568 propLet
747592 function vbH(){try{return chrome.webview.hostObjects.sync.vb_Host} catch(e){}}
747756 function RunFunc(funcName, token) {
747832   var ctx  = window;
747880   var args = Array.prototype.slice.call(arguments, 2);
747996 DispCallFunc
748076   for(var i= 0; i < ns.length; i++){ ctx = ctx[ns[i]] }
748192   var eMsg = '', res = null;
748256   try{res  = ctx[func].apply(ctx, args)} catch(e){eMsg='Err in '+func+'(): '+e.message};
748440   try{vbH().RaiseResultEvent(res, token, eMsg);} catch(e){}
748564 function Cbn(obj, callType, methodName, token) {
748668   try{switch (callType) {
748724     case 1: res = obj[methodName].apply(obj, Array.prototype.slice.call(arguments, 4));break;
748916 function GetProp(propName, token){
748992 contextmenu
749040     case 2: try{vbH().RaiseMessageEvent (methodName, JSON.stringify(obj)); res = obj[methodName]} catch(e){};break;
749276     case 4: res = obj[methodName]=arguments[4];break;
749388   }} catch(e){eMsg='Err in '+methodName+'(): '+e.message};
749512 propGet
749536   var ctx = window, ns   = propName.split('.'), prop = ns.pop();
749672   try{res = ctx[prop]} catch(e){eMsg='Err in '+propName+'-Get: '+e.message};
749832 function LetProp(propName, RHS, token){
749916   try{ctx[prop]=RHS} catch(e){eMsg=e.message};
750028   try{vbH().RaiseResultEvent(true, token, eMsg);} catch(e){}
750156 document.addEventListener('contextmenu', function(e){
750268   try{vbH().RaiseMessageEvent('contextmenu', JSON.stringify({x:e.screenX, y:e.screenY})) }catch(e){}
750488 onblur
750508 window.onblur =  function(){ try{vbH().RaiseMessageEvent('onblur', '')} catch(e){} }
750684 window.onfocus = function(){ try{vbH().RaiseMessageEvent('onfocus','')} catch(e){} }
750984 try{cbFuncs={run:RunFunc, cbn:Cbn, propGet:GetProp, propLet:LetProp}; vbH().SetFuncObj(cbFuncs);} catch(e){}
```

Notes:

- `vbH()` resolves the host object registered as `vb_Host` (RC6's name for
  what our port registers as `vbHost`).
- `RaiseResultEvent(res, token, eMsg)` is the return channel: it is a
  NESTED synchronous host-object call made while the VB6 side still sits
  inside its outgoing `run`/`cbn`/`propGet`/`propLet` dispatch.
- The `contextmenu`/`onblur`/`onfocus` handlers show RC6 raises those
  through reserved `RaiseMessageEvent` message names - it has no dedicated
  `RaiseContextMenuEvent` (that method is our port's addition).

## 2. AddObject window-alias glue (offsets 747204-747292)

```
747204 try{
747220 =chrome.webview.hostObjects.sync.
747292 }catch(e){}
```

Concatenation skeleton for the `AddObject` alias script, i.e.
`try{ <Name>=chrome.webview.hostObjects.sync.<Name> }catch(e){}` -
equivalent to our port's `window[...] = chrome.webview.hostObjects.sync[...]`.

## 3. The VB6-side dispatch machinery (offsets 742280-746220 + imports)

```
742280 we need 12 ParamArray-entries (in addition to the PropertyName)
745164 pVTableToUse can't be a Null-Pointer
745244 pVarPtrNewInstance can't be a Null-Pointer
745400 we need a valid Implementing-Instance to call-back-to
745528 pVTableDst can't be a Null-Pointer
745812 Invalid IID-String
746220 currently only up to 10 Arguments are supported by vbIDispatch->Invoke
751492  TimeOut, waiting for Result
```

ANSI (import/thunk area):

```
494808 DispCallFunc          <- oleaut32!DispCallFunc import (VB6 Declare)
```

Typelib symbol (Unicode, 3584908): `_IID_VBIDISPATCH`.

The corresponding public interface from RC6.idl (line 12489):

```idl
interface _vbIDispatch : IDispatch {
    HRESULT GetIDForMemberName(
                    [in, out] long* UserData,
                    [in] long pVTable,
                    [in, out] BSTR* MemberName,
                    [out, retval] long* );
    [vararg]
    HRESULT Invoke(
                    [in, out] long* UserData,
                    [in] long pVTable,
                    [in] long DispID,
                    [in] VbCallType CallType,
                    [in, out] VARIANT* VResult,
                    [in, out] SAFEARRAY(VARIANT)* P,
                    [out, retval] hResult* );
};
coclass vbIDispatch { [default] interface _vbIDispatch; };  // 73BF6C24-B56D-4333-B9D1-B7B12762956B
```

Interpretation:

- RC6 ships a general-purpose raw-dispatch helper (`vbIDispatch`) that
  takes a NAKED interface pointer (`pVTable` as long), resolves a DISPID
  from a member name, and performs `IDispatch::Invoke` with hand-built
  DISPPARAMS - implemented over the `DispCallFunc` import, capped at 10
  arguments ("currently only up to 10 Arguments are supported by
  vbIDispatch->Invoke").
- This is how `jsRun`/`jsProp`/`jsCallByName` call into the `cbFuncs`
  proxy received via `SetFuncObj`: raw vtable `GetIDsOfNames` + `Invoke`,
  NOT VB6 late binding (which fails on these proxies with
  DISP_E_UNKNOWNLCID) and NOT IDispatchEx (the proxy does not QI to it).
- " TimeOut, waiting for Result" is RC6's jsRun timeout message
  (`jsCallTimeOutSeconds`) - a pump-with-timeout guards the wait for the
  token result in case the nested `RaiseResultEvent` did not arrive
  synchronously.

## 4. Assorted cWebView2-related strings (same region)

```
746884 \WebView2Loader.dll
746928 \Microsoft\
746956 \Application
746988 WebView2Loader.dll not found
747052 HosthWnd may not be 0
747100 location.href
747136 document.title
751224 HasPosition
751252 HasSize
751272 ShouldDisplayMenuBar
751320 ShouldDisplayScrollBars
751372 ShouldDisplayStatus
751416 ShouldDisplayToolbar
750924 Blocked
```

- `\Microsoft\` + `\Application` are the building blocks of
  `GetMostRecentInstallPath` (matches our reimplementation).
- `HosthWnd may not be 0` is the error RC6 raises from `BindTo(0)`.
- The `HasPosition`..`ShouldDisplayToolbar` keys are the
  `NewWindowFeatures` collection keys (match our port).

## 5. Open question for our pvBridgeInvoke

Our port replicates the same architecture (bridge JS + raw vtable
`GetIDsOfNames`/`Invoke` with manual DISPPARAMS), but on the current
Evergreen runtime the func-object proxy's `GetIDsOfNames` returned
DISP_E_UNKNOWNLCID for LCIDs 0, user-default and &H409, and QI for
IDispatchEx fails (err 13). Since RC6 demonstrably succeeds through the
same entry points, the difference must be in a call detail - the LCID
value being the prime suspect (candidates not yet tried: &H400
LOCALE_USER_DEFAULT, &H800 LOCALE_SYSTEM_DEFAULT, &H7F LOCALE_INVARIANT)
or a DISPPARAMS/argument-marshaling detail.
