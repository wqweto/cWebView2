VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' CLASS: cWebView2
' RC6-compatible replacement of vbRichClient6's cWebView2, built directly
' on the WebView2 COM API as internalized VB6-safe interfaces in
' VBWebView2Impl.tlb - no RC6.dll and no dependency on the original
' WebView2 SDK typelib. Targets the Evergreen runtime via WebView2Loader.dll
' deployed next to the exe. Blocking calls (BindTo, Navigate, jsRun,
' CapturePreview etc) spin the thread message pump until the native async
' completion arrives or the timeout elapses. All native callbacks land in
' cWebView2Callback instances which forward to the Friend sinks here
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Public enums
'=========================================================================

Public Enum eWebView2FocusReason
    FocusReason_PROGRAMMATIC = 0
    FocusReason_NEXT = 1
    FocusReason_PREVIOUS = 2
End Enum

Public Enum eWebView2ImageCaptureFormat
    CaptureAs_PNG = 0
    CaptureAs_JPG = 1
End Enum

Public Enum eWebView2PermissionKind
    PERMISSION_KIND_UNKNOWN_PERMISSION = 0
    PERMISSION_KIND_MICROPHONE = 1
    PERMISSION_KIND_CAMERA = 2
    PERMISSION_KIND_GEOLOCATION = 3
    PERMISSION_KIND_NOTIFICATIONS = 4
    PERMISSION_KIND_OTHER_SENSORS = 5
    PERMISSION_KIND_CLIPBOARD_READ = 6
End Enum

Public Enum eWebView2PermissionState
    PERMISSION_STATE_DEFAULT = 0
    PERMISSION_STATE_ALLOW = 1
    PERMISSION_STATE_DENY = 2
End Enum

Public Enum eWebView2ScriptDialogKind
    SCRIPT_DIALOG_KIND_ALERT = 0
    SCRIPT_DIALOG_KIND_CONFIRM = 1
    SCRIPT_DIALOG_KIND_PROMPT = 2
    SCRIPT_DIALOG_KIND_BEFOREUNLOAD = 3
End Enum

Public Enum eWebView2ResourceFilter
    Filter_ALL = 0
    Filter_DOCUMENT = 1
    Filter_STYLESHEET = 2
    Filter_IMAGE = 3
    Filter_MEDIA = 4
    Filter_FONT = 5
    Filter_SCRIPT = 6
    Filter_XML_HTTP_REQUEST = 7
    Filter_FETCH = 8
    Filter_TEXT_TRACK = 9
    Filter_EVENT_SOURCE = 10
    Filter_WEBSOCKET = 11
    Filter_MANIFEST = 12
    Filter_SIGNED_EXCHANGE = 13
    Filter_PING = 14
    Filter_CSP_VIOLATION_REPORT = 15
    Filter_OTHER = 16
End Enum

Public Enum eWebView2AccKeyState
    Key_DOWN = 0
    Key_UP = 1
    SysKey_DOWN = 2
    SysKey_UP = 3
End Enum

'=========================================================================
' Public events
'=========================================================================

Public Event InitComplete()
Public Event DocumentComplete()
Public Event ProcessFailed()
Public Event SourceChanged()
Public Event NavigationStarting(ByVal IsUserInitiated As Boolean, ByVal IsRedirected As Boolean, ByVal URI As String, Cancel As Boolean)
Public Event NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
Public Event UserContextMenu(ByVal ScreenX As Long, ByVal ScreenY As Long)
Public Event JSMessage(ByVal sMsg As String, ByVal sMsgContent As String, oJSONContent As Collection)
Public Event JSAsyncResult(Result As Variant, ByVal Token As Currency, ByVal ErrString As String)
Public Event GotFocus(ByVal Reason As eWebView2FocusReason)
Public Event LostFocus(ByVal Reason As eWebView2FocusReason)
Public Event AcceleratorKeyPressed(ByVal KeyState As eWebView2AccKeyState, ByVal IsExtendedKey As Boolean, ByVal WasKeyDown As Boolean, ByVal IsKeyReleased As Boolean, ByVal IsMenuKeyDown As Boolean, ByVal RepeatCount As Long, ByVal ScanCode As Long, IsHandled As Boolean)
Public Event NewWindowRequested(ByVal IsUserInitiated As Boolean, IsHandled As Boolean, ByVal URI As String, NewWindowFeatures As Collection)
Public Event PermissionRequested(ByVal IsUserInitiated As Boolean, State As eWebView2PermissionState, ByVal URI As String, ByVal PermissionKind As eWebView2PermissionKind)
Public Event ScriptDialogOpening(ByVal ScriptDialogKind As eWebView2ScriptDialogKind, Accept As Boolean, ResultText As String, ByVal URI As String, ByVal Message As String, ByVal DefaultText As String)
Public Event WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, RespHeaders As Collection)
Public Event FrameCreated(ByVal frameName As String)
Public Event DownloadStarting(ByVal URI As String, ByVal MimeType As String, ByVal LocalSuggestedPath As String, ByVal LocalSuggestedFile As String, CancelBrowserDialog As Boolean, CancelDownloadOp As Boolean, NewLocalFilePath As String)
Public Event WebResourceRequested(ByVal URI As String, ByVal FilterType As eWebView2ResourceFilter, Permit As eWebView2PermissionState)
Public Event TitleChange(ByVal Text As String)
Public Event ContentLoading(ByVal IsErrorPage As Boolean)
Public Event HistoryChanged()
Public Event WindowCloseRequested()
Public Event ZoomFactorChanged()
Public Event MoveFocusRequested(ByVal Reason As eWebView2FocusReason, Handled As Boolean)
Public Event ContextMenuRequested(ByVal PageURI As String, ByVal LinkURI As String, ByVal SelectionText As String, ByVal ScreenX As Long, ByVal ScreenY As Long, Handled As Boolean)

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As APIRECT) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, ByVal lpsz As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader" (ByVal browserExecutableFolder As Long, ByVal userDataFolder As Long, ByVal environmentOptions As Long, ByVal environmentCreatedHandler As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CoWaitForMultipleHandles Lib "ole32" (ByVal dwFlags As Long, ByVal dwTimeout As Long, ByVal cHandles As Long, pHandles As Any, lpdwindex As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByVal lpMsg As Long, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (ByVal lpMsg As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (ByVal lpMsg As Long) As Long
Private Declare Sub VariantClear Lib "oleaut32" (pvarg As Variant)
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As IVBStream) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long

Private Type APIMSG
    hWnd                As Long
    lMessage            As Long
    wParam              As Long
    lParam              As Long
    time                As Long
    pt                  As APIPOINT
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oEnvironment              As IVBCoreWebView2Environment
Private m_oController               As IVBCoreWebView2Controller
Private m_oWebView                  As IVBCoreWebView2
Private m_oWebView4                 As IVBCoreWebView2_4
Private m_oWebView11                As IVBCoreWebView2_11
Private m_oWebView13                As IVBCoreWebView2_13
Private m_oSettings                 As IVBCoreWebView2Settings
Private m_oSettings6                As IVBCoreWebView2Settings6
Private m_oEventSink                As cWebView2Callback
Private m_oFocusGotSink             As cWebView2Callback
Private m_oFocusLostSink            As cWebView2Callback
Private m_eLastFocusReason          As eWebView2FocusReason
Private m_hWndHost                  As Long
Private m_bInitComplete             As Boolean
Private m_hInitError                As Long
Private m_bNavDone                  As Boolean
Private m_bNavSuccess               As Boolean
Private m_lNavErrorStatus           As Long
Private m_bShuttingDown             As Boolean
Private m_dbljsCallTimeOutSeconds   As Double
Private m_cTokenSeq                 As Currency
'--- keys are prefixed by entry kind: "T" & token - completed result, "P" &
'--- token - sync waiter pending, "A" & token - async result expected
Private m_cScriptResults            As Collection
Private m_sJsLastError              As String

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & sFunction & "]"
End Sub

Private Sub pvNotImplemented(sFunction As String)
    Err.Raise vbObjectError, "cWebView2", sFunction & " is not yet implemented"
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get CanGoBack() As Boolean
    If Not m_oWebView Is Nothing Then
        CanGoBack = m_oWebView.CanGoBack
    End If
End Property

Public Property Get CanGoForward() As Boolean
    If Not m_oWebView Is Nothing Then
        CanGoForward = m_oWebView.CanGoForward
    End If
End Property

Public Property Get DocumentURL() As String
    If Not m_oWebView Is Nothing Then
        DocumentURL = pvToString(m_oWebView.Source)
    End If
End Property

Public Property Get DocumentTitle() As String
    If Not m_oWebView Is Nothing Then
        DocumentTitle = pvToString(m_oWebView.DocumentTitle)
    End If
End Property

Public Property Get BrowserProcessId() As Long
    If Not m_oWebView Is Nothing Then
        BrowserProcessId = m_oWebView.BrowserProcessId
    End If
End Property

Public Property Get HosthWnd() As Long
    HosthWnd = m_hWndHost
End Property

Public Property Get ZoomFactor() As Double
    If Not m_oController Is Nothing Then
        ZoomFactor = m_oController.ZoomFactor
    End If
End Property

Public Property Let ZoomFactor(ByVal dblValue As Double)
    If Not m_oController Is Nothing Then
        m_oController.ZoomFactor = dblValue
    End If
End Property

Public Property Get IsZoomControlEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsZoomControlEnabled = m_oSettings.IsZoomControlEnabled
    End If
End Property

Public Property Let IsZoomControlEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsZoomControlEnabled = -bValue
    End If
End Property

Public Property Get IsScriptEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsScriptEnabled = m_oSettings.IsScriptEnabled
    End If
End Property

Public Property Let IsScriptEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsScriptEnabled = -bValue
    End If
End Property

Public Property Get IsWebMessageEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsWebMessageEnabled = m_oSettings.IsWebMessageEnabled
    End If
End Property

Public Property Let IsWebMessageEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsWebMessageEnabled = -bValue
    End If
End Property

Public Property Get AreDefaultScriptDialogsEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        AreDefaultScriptDialogsEnabled = m_oSettings.AreDefaultScriptDialogsEnabled
    End If
End Property

Public Property Let AreDefaultScriptDialogsEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreDefaultScriptDialogsEnabled = -bValue
    End If
End Property

Public Property Get IsStatusBarEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsStatusBarEnabled = m_oSettings.IsStatusBarEnabled
    End If
End Property

Public Property Let IsStatusBarEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsStatusBarEnabled = -bValue
    End If
End Property

Public Property Get AreDevToolsEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        AreDevToolsEnabled = m_oSettings.AreDevToolsEnabled
    End If
End Property

Public Property Let AreDevToolsEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreDevToolsEnabled = -bValue
    End If
End Property

Public Property Get AreDefaultContextMenusEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        AreDefaultContextMenusEnabled = m_oSettings.AreDefaultContextMenusEnabled
    End If
End Property

Public Property Let AreDefaultContextMenusEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreDefaultContextMenusEnabled = -bValue
    End If
End Property

Public Property Get AreHostObjectsAllowed() As Boolean
    If Not m_oSettings Is Nothing Then
        AreHostObjectsAllowed = m_oSettings.AreHostObjectsAllowed
    End If
End Property

Public Property Let AreHostObjectsAllowed(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.AreHostObjectsAllowed = -bValue
    End If
End Property

Public Property Get jsCallTimeOutSeconds() As Double
    jsCallTimeOutSeconds = m_dbljsCallTimeOutSeconds
End Property

Public Property Let jsCallTimeOutSeconds(ByVal dblValue As Double)
    m_dbljsCallTimeOutSeconds = dblValue
End Property

Public Property Get jsProp(ByVal PropName As String) As Variant
    pvExecScriptSync PropName, jsProp
End Property

Public Property Let jsProp(ByVal PropName As String, ByVal RHS As Variant)
    Dim vResult         As Variant

    pvExecScriptSync PropName & "=" & pvJsonEncode(RHS), vResult
End Property

Public Property Get jsLastError() As String
    jsLastError = m_sJsLastError
End Property

Public Property Get UserAgent() As String
    If Not m_oSettings6 Is Nothing Then
        UserAgent = pvToString(m_oSettings6.UserAgent)
    End If
End Property

Public Property Let UserAgent(ByVal sValue As String)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.UserAgent = StrPtr(sValue)
    End If
End Property

Public Property Get AreBrowserAcceleratorKeysEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        AreBrowserAcceleratorKeysEnabled = m_oSettings6.AreBrowserAcceleratorKeysEnabled
    End If
End Property

Public Property Let AreBrowserAcceleratorKeysEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.AreBrowserAcceleratorKeysEnabled = -bValue
    End If
End Property

Public Property Get IsPasswordAutosaveEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsPasswordAutosaveEnabled = m_oSettings6.IsPasswordAutosaveEnabled
    End If
End Property

Public Property Let IsPasswordAutosaveEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsPasswordAutosaveEnabled = -bValue
    End If
End Property

Public Property Get IsGeneralAutofillEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsGeneralAutofillEnabled = m_oSettings6.IsGeneralAutofillEnabled
    End If
End Property

Public Property Let IsGeneralAutofillEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsGeneralAutofillEnabled = -bValue
    End If
End Property

Public Property Get IsPinchZoomEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsPinchZoomEnabled = m_oSettings6.IsPinchZoomEnabled
    End If
End Property

Public Property Let IsPinchZoomEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsPinchZoomEnabled = -bValue
    End If
End Property

Public Property Get IsSwipeNavigationEnabled() As Boolean
    If Not m_oSettings6 Is Nothing Then
        IsSwipeNavigationEnabled = m_oSettings6.IsSwipeNavigationEnabled
    End If
End Property

Public Property Let IsSwipeNavigationEnabled(ByVal bValue As Boolean)
    If Not m_oSettings6 Is Nothing Then
        m_oSettings6.IsSwipeNavigationEnabled = -bValue
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function BindTo( _
        ByVal HosthWnd As Long, _
        Optional ByVal SecondsToWaitForInitComplete As Double = 8, _
        Optional ByVal browserInstallPath As String, _
        Optional ByVal userDataFolder As String, _
        Optional ByVal AdditionalBrowserArguments As String, _
        Optional ByVal AllowSingleSignOnUsingOSPrimaryAccount As Boolean, _
        Optional ByVal ExclusiveUserDataFolderAccess As Boolean, _
        Optional ByVal IsCustomCrashReportingEnabled As Boolean, _
        Optional ByVal EnableTrackingPrevention As Boolean = True, _
        Optional ByVal AreBrowserExtensionsEnabled As Boolean) As Long
    Const RPC_E_TIMEOUT                 As Long = &H8001011F
    Dim oOptions        As cWebView2Callback
    Dim oHandler        As cWebView2Callback
    Dim hResult         As Long

    On Error GoTo EH
    m_hWndHost = HosthWnd
    '--- allow re-binding after a previous Shutdown on this instance
    m_bShuttingDown = False
    m_bInitComplete = False
    m_hInitError = 0
    Set oOptions = New cWebView2Callback
    oOptions.AdditionalBrowserArguments = AdditionalBrowserArguments
    oOptions.AllowSingleSignOnUsingOSPrimaryAccount = AllowSingleSignOnUsingOSPrimaryAccount
    oOptions.ExclusiveUserDataFolderAccess = ExclusiveUserDataFolderAccess
    oOptions.IsCustomCrashReportingEnabled = IsCustomCrashReportingEnabled
    oOptions.EnableTrackingPrevention = EnableTrackingPrevention
    oOptions.AreBrowserExtensionsEnabled = AreBrowserExtensionsEnabled
    Set oHandler = New cWebView2Callback
    oHandler.Init Me
    hResult = pvCreateWebView2Environment(browserInstallPath, userDataFolder, oOptions, oHandler)
    If hResult <> 0 Then
        BindTo = hResult
        Exit Function
    End If
    If Not pvPumpUntil(m_bInitComplete, SecondsToWaitForInitComplete) Then
        BindTo = RPC_E_TIMEOUT
        Exit Function
    End If
    BindTo = m_hInitError
    Exit Function
EH:
    BindTo = Err.Number
    PrintError "BindTo"
End Function

Public Sub SyncSizeToHostWindow()
    Dim uRect           As APIRECT

    On Error GoTo EH
    Call GetClientRect(m_hWndHost, uRect)
    If Not m_oController Is Nothing Then
        m_oController.SetBounds 0, 0, uRect.Right, uRect.Bottom
    End If
    Exit Sub
EH:
    PrintError "SyncSizeToHostWindow"
End Sub

Public Sub MoveRelativeToHostWindow(x As Variant, y As Variant, dx As Variant, dy As Variant)
    Dim uRect           As APIRECT
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long

    On Error GoTo EH
    uRect = m_oController.Bounds
    lLeft = uRect.Left
    lTop = uRect.Top
    lWidth = uRect.Right - uRect.Left
    lHeight = uRect.Bottom - uRect.Top
    If Not IsMissing(x) Then
        lLeft = CLng(x)
    End If
    If Not IsMissing(y) Then
        lTop = CLng(y)
    End If
    If Not IsMissing(dx) Then
        lWidth = CLng(dx)
    End If
    If Not IsMissing(dy) Then
        lHeight = CLng(dy)
    End If
    m_oController.SetBounds lLeft, lTop, lLeft + lWidth, lTop + lHeight
    x = lLeft
    y = lTop
    dx = lWidth
    dy = lHeight
    Exit Sub
EH:
    PrintError "MoveRelativeToHostWindow"
End Sub

Public Sub SetFocus(Optional ByVal Reason As eWebView2FocusReason = FocusReason_PROGRAMMATIC)
    On Error GoTo EH
    m_eLastFocusReason = Reason
    If Not m_oController Is Nothing Then
        m_oController.MoveFocus Reason
    End If
    Exit Sub
EH:
    PrintError "SetFocus"
End Sub

Public Function Navigate(ByVal URI As String, Optional ByVal SecondsToWaitForDocumentComplete As Double = 8) As Variant
    On Error GoTo EH
    m_bNavDone = False
    m_bNavSuccess = True
    m_oWebView.Navigate StrPtr(URI)
    If Not pvPumpUntil(m_bNavDone, SecondsToWaitForDocumentComplete) Then
        Navigate = Empty
        Exit Function
    End If
    Navigate = m_bNavSuccess
    Exit Function
EH:
    Navigate = Empty
    PrintError "Navigate"
End Function

Public Function NavigateToString(ByVal sHTMLContent As String, Optional ByVal SecondsToWaitForDocumentComplete As Double = 8) As Variant
    On Error GoTo EH
    m_bNavDone = False
    m_bNavSuccess = True
    m_oWebView.NavigateToString StrPtr(sHTMLContent)
    If Not pvPumpUntil(m_bNavDone, SecondsToWaitForDocumentComplete) Then
        Exit Function
    End If
    NavigateToString = m_bNavSuccess
    Exit Function
EH:
    PrintError "NavigateToString"
End Function

Public Sub Reload()
    If Not m_oWebView Is Nothing Then
        m_oWebView.Reload
    End If
End Sub

Public Sub GoBack()
    If Not m_oWebView Is Nothing Then
        m_oWebView.GoBack
    End If
End Sub

Public Sub GoForward()
    If Not m_oWebView Is Nothing Then
        m_oWebView.GoForward
    End If
End Sub

Public Sub OpenDevToolsWindow()
    If Not m_oWebView Is Nothing Then
        m_oWebView.OpenDevToolsWindow
    End If
End Sub

Public Function GetMostRecentInstallPath(Optional ByVal EdgeChannel As String = "EdgeWebView", Optional VersionString As String) As String
    #If EdgeChannel And VersionString Then
    #End If
    pvNotImplemented "GetMostRecentInstallPath"
End Function

Public Function NavigateWithWebResourceRequest(ByVal URI As String, Optional ByVal Method As String = "GET", Optional ByVal Headers_CrLfSeparated As String, Optional ContentStream As Object, Optional ByVal SecondsToWaitForRequestComplete As Double = 8) As Variant
    Dim oEnv2           As IVBCoreWebView2Environment2
    Dim oRequest        As IVBCoreWebView2WebResourceRequest
    Dim oStream         As IVBStream

    On Error GoTo EH
    Set oEnv2 = m_oEnvironment
    If Not ContentStream Is Nothing Then
        Set oStream = ContentStream
    End If
    Set oRequest = oEnv2.CreateWebResourceRequest(StrPtr(URI), StrPtr(Method), oStream, StrPtr(Headers_CrLfSeparated))
    m_bNavDone = False
    m_bNavSuccess = True
    m_oWebView4.NavigateWithWebResourceRequest oRequest
    If pvPumpUntil(m_bNavDone, SecondsToWaitForRequestComplete) Then
        NavigateWithWebResourceRequest = m_bNavSuccess
    End If
    Exit Function
EH:
    PrintError "NavigateWithWebResourceRequest"
End Function

Public Sub AddWebResourceRequestedFilter(ByVal sFilter As String, Optional ByVal FilterContext As eWebView2ResourceFilter = Filter_ALL)
    On Error GoTo EH
    m_oWebView.AddWebResourceRequestedFilter StrPtr(sFilter), FilterContext
    Exit Sub
EH:
    PrintError "AddWebResourceRequestedFilter"
End Sub

Public Sub RemoveWebResourceRequestedFilter(ByVal sFilter As String, ByVal FilterContext As eWebView2ResourceFilter)
    On Error GoTo EH
    m_oWebView.RemoveWebResourceRequestedFilter StrPtr(sFilter), FilterContext
    Exit Sub
EH:
    PrintError "RemoveWebResourceRequestedFilter"
End Sub

Public Sub AddScriptToExecuteOnDocumentCreated(ByVal jsCode As String)
    Dim oHandler        As cWebView2Callback

    On Error GoTo EH
    Set oHandler = New cWebView2Callback
    oHandler.Init Me
    m_oWebView.AddScriptToExecuteOnDocumentCreated StrPtr(jsCode), oHandler
    Exit Sub
EH:
    PrintError "AddScriptToExecuteOnDocumentCreated"
End Sub

Public Sub ExecuteScript(ByVal jsCode As String)
    Dim oHandler        As cWebView2Callback

    On Error GoTo EH
    Set oHandler = New cWebView2Callback
    oHandler.Init Me
    m_oWebView.ExecuteScript StrPtr(jsCode), oHandler
    Exit Sub
EH:
    PrintError "ExecuteScript"
End Sub

Public Sub AddObject(ByVal ObjName As String, Obj As Variant)
    Dim sAlias          As String

    On Error GoTo EH
    m_oWebView.AddHostObjectToScript StrPtr(ObjName), Obj
    '--- RC6 compat: publish a global window.<name> alias of the sync proxy so
    '--- page script can call <name>.MethodName(...) directly - both in the
    '--- current document and in all future ones
    sAlias = "window[" & pvJsonEncodeString(ObjName) & "] = chrome.webview.hostObjects.sync[" & pvJsonEncodeString(ObjName) & "];"
    ExecuteScript sAlias
    AddScriptToExecuteOnDocumentCreated sAlias
    Exit Sub
EH:
    PrintError "AddObject"
End Sub

Public Sub RemoveObject(ByVal ObjName As String)
    On Error GoTo EH
    m_oWebView.RemoveHostObjectFromScript StrPtr(ObjName)
    ExecuteScript "delete window[" & pvJsonEncodeString(ObjName) & "];"
    Exit Sub
EH:
    PrintError "RemoveObject"
End Sub

Public Sub RaiseMessageEvent(sMsg As Variant, sMsgContent As Variant)
    Dim sContent        As String
    Dim vResult         As Variant
    Dim oContent        As Collection
    Dim lIdx            As Long

    On Error GoTo EH
    sContent = sMsgContent
    '--- populate oJSONContent only when the content looks like a JSON object/array
    Select Case Left$(LTrim$(sContent), 1)
    Case "{", "["
        lIdx = 1
        pvJsonParseValue sContent, lIdx, vResult
        If IsObject(vResult) Then
            Set oContent = vResult
        End If
    End Select
    RaiseEvent JSMessage(CStr(sMsg), sContent, oContent)
    Exit Sub
EH:
    PrintError "RaiseMessageEvent"
End Sub

Public Sub RaiseResultEvent(Res As Variant, ByVal Token As Variant, ByVal ErrMsg As Variant)
    On Error GoTo EH
    RaiseEvent JSAsyncResult(Res, CCur(Token), CStr(ErrMsg))
    Exit Sub
EH:
    PrintError "RaiseResultEvent"
End Sub

Public Sub RaiseContextMenuEvent(ByVal ScreenX As Variant, ByVal ScreenY As Variant)
    On Error GoTo EH
    RaiseEvent UserContextMenu(CLng(ScreenX), CLng(ScreenY))
    Exit Sub
EH:
    PrintError "RaiseContextMenuEvent"
End Sub

Public Function CallDevToolsProtocolMethod(ByVal MethodName As String, ByVal ParamsAsJSON As String) As Variant
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant
    Dim sJson           As String
    Dim lIdx            As Long

    On Error GoTo EH
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    m_oWebView.CallDevToolsProtocolMethod StrPtr(MethodName), StrPtr(ParamsAsJSON), oHandler
    '--- jsCallTimeOutSeconds 0 (or less) means fire-and-forget: CDP response
    '--- is discarded on arrival
    If m_dbljsCallTimeOutSeconds <= 0 Then
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, m_dbljsCallTimeOutSeconds) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then
        Exit Function
    End If
    sJson = vItem(1)
    lIdx = 1
    pvJsonParseValue sJson, lIdx, CallDevToolsProtocolMethod
    Exit Function
EH:
    PrintError "CallDevToolsProtocolMethod"
End Function

Public Function CapturePreview(Optional ByVal ImageFormat As eWebView2ImageCaptureFormat = CaptureAs_PNG) As Byte()
    Const STREAM_SEEK_SET               As Long = 0
    Const STREAM_SEEK_END               As Long = 2
    Dim oHandler        As cWebView2Callback
    Dim oStream         As IVBStream
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant
    Dim cPos            As Currency
    Dim lSize           As Long
    Dim lRead           As Long
    Dim baResult()      As Byte

    On Error GoTo EH
    If CreateStreamOnHGlobal(0, 1, oStream) < 0 Then
        Exit Function
    End If
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    m_oWebView.CapturePreview ImageFormat, oStream, oHandler
    If m_dbljsCallTimeOutSeconds <= 0 Then
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, m_dbljsCallTimeOutSeconds) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then
        Exit Function
    End If
    '--- IVBStream positions travel as CURRENCY i.e. scaled by 1/10000
    oStream.Seek 0, STREAM_SEEK_END, cPos
    lSize = cPos * 10000
    If lSize > 0 Then
        oStream.Seek 0, STREAM_SEEK_SET, cPos
        ReDim baResult(0 To lSize - 1)
        oStream.Read VarPtr(baResult(0)), lSize, lRead
        CapturePreview = baResult
    End If
    Exit Function
EH:
    PrintError "CapturePreview"
End Function

Public Function PrintToPdf( _
        ByVal ResultFilePath As String, _
        Optional ByVal SecondsToWaitForPrintComplete As Double = 8, _
        Optional ByVal Landscape As Boolean, _
        Optional ByVal ScaleFactor As Double, _
        Optional ByVal PageWidth As Double, _
        Optional ByVal PageHeight As Double, _
        Optional ByVal MarginTop As Double = -1, _
        Optional ByVal MarginBottom As Double = -1, _
        Optional ByVal MarginLeft As Double = -1, _
        Optional ByVal MarginRight As Double = -1, _
        Optional ByVal ShouldPrintBackgrounds As Boolean, _
        Optional ByVal ShouldPrintSelectionOnly As Boolean, _
        Optional ByVal ShouldPrintHeaderAndFooter As Boolean = True, _
        Optional ByVal HeaderTitle As String, _
        Optional ByVal FooterUri As String) As Variant
    Const ORIENT_LANDSCAPE              As Long = 1
    Dim oEnv6           As IVBCoreWebView2Environment6
    Dim oSettings       As IVBCoreWebView2PrintSettings
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView11 Is Nothing Then
        Exit Function
    End If
    On Error Resume Next
    Set oEnv6 = m_oEnvironment
    On Error GoTo EH
    If Not oEnv6 Is Nothing Then
        Set oSettings = oEnv6.CreatePrintSettings()
        If Landscape Then
            oSettings.Orientation = ORIENT_LANDSCAPE
        End If
        If ScaleFactor > 0 Then
            oSettings.ScaleFactor = ScaleFactor
        End If
        If PageWidth > 0 Then
            oSettings.PageWidth = PageWidth
        End If
        If PageHeight > 0 Then
            oSettings.PageHeight = PageHeight
        End If
        If MarginTop >= 0 Then
            oSettings.MarginTop = MarginTop
        End If
        If MarginBottom >= 0 Then
            oSettings.MarginBottom = MarginBottom
        End If
        If MarginLeft >= 0 Then
            oSettings.MarginLeft = MarginLeft
        End If
        If MarginRight >= 0 Then
            oSettings.MarginRight = MarginRight
        End If
        oSettings.ShouldPrintBackgrounds = -ShouldPrintBackgrounds
        oSettings.ShouldPrintSelectionOnly = -ShouldPrintSelectionOnly
        oSettings.ShouldPrintHeaderAndFooter = -ShouldPrintHeaderAndFooter
        If LenB(HeaderTitle) <> 0 Then
            oSettings.HeaderTitle = StrPtr(HeaderTitle)
        End If
        If LenB(FooterUri) <> 0 Then
            oSettings.FooterUri = StrPtr(FooterUri)
        End If
    End If
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    m_oWebView11.PrintToPdf StrPtr(ResultFilePath), oSettings, oHandler
    If SecondsToWaitForPrintComplete <= 0 Then
        PrintToPdf = True
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsToWaitForPrintComplete) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then
        Exit Function
    End If
    PrintToPdf = (vItem(1) = "true")
    Exit Function
EH:
    PrintError "PrintToPdf"
End Function

Public Function AddBrowserExtension(ByVal ExtensionFolderPath As String, Optional ByVal SecondsToWaitForComplete As Double = 8) As String
    Dim oProfile        As IVBCoreWebView2Profile7
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    Set oProfile = pvGetProfile7()
    If oProfile Is Nothing Then
        Exit Function
    End If
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    oProfile.AddBrowserExtension StrPtr(ExtensionFolderPath), oHandler
    If SecondsToWaitForComplete <= 0 Then
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsToWaitForComplete) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then
        Exit Function
    End If
    AddBrowserExtension = vItem(1)
    Exit Function
EH:
    PrintError "AddBrowserExtension"
End Function

Public Function GetBrowserExtensions(Optional ByVal SecondsToWaitForComplete As Double = 8) As Collection
    Dim oProfile        As IVBCoreWebView2Profile7
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    Set oProfile = pvGetProfile7()
    If oProfile Is Nothing Then
        Exit Function
    End If
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    oProfile.GetBrowserExtensions oHandler
    If SecondsToWaitForComplete <= 0 Then
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, SecondsToWaitForComplete) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then
        Exit Function
    End If
    Set GetBrowserExtensions = vItem(1)
    Exit Function
EH:
    PrintError "GetBrowserExtensions"
End Function

Public Function CallByName(ByVal Obj As Object, ByVal ProcName As String, ByVal CallType As VbCallType, ParamArray P() As Variant) As Variant
    Select Case UBound(P)
    Case -1
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType)
    Case 0
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0))
    Case 1
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0), P(1))
    Case 2
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0), P(1), P(2))
    Case 3
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0), P(1), P(2), P(3))
    Case 4
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0), P(1), P(2), P(3), P(4))
    Case 5
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0), P(1), P(2), P(3), P(4), P(5))
    Case 6
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0), P(1), P(2), P(3), P(4), P(5), P(6))
    Case 7
        pvAssignVariant CallByName, VBA.CallByName(Obj, ProcName, CallType, P(0), P(1), P(2), P(3), P(4), P(5), P(6), P(7))
    Case Else
        Err.Raise 450                               '--- wrong number of arguments
    End Select
End Function

Public Function jsCallByName(ByVal Obj As Object, ByVal MethodName As String, ByVal CallType As VbCallType, ParamArray P() As Variant) As Variant
    #If Obj And MethodName And CallType And P Then
    #End If
    pvNotImplemented "jsCallByName"
End Function

Public Function jsRun(ByVal FuncName As String, ParamArray P() As Variant) As Variant
    Dim vArgs           As Variant

    vArgs = P
    pvExecScriptSync pvJsBuildCall(FuncName, vArgs), jsRun
End Function

Public Function jsRunAsync(ByVal FuncName As String, ParamArray P() As Variant) As Currency
    Dim vArgs           As Variant
    Dim oHandler        As cWebView2Callback

    On Error GoTo EH
    vArgs = P
    jsRunAsync = pvNextToken()
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, jsRunAsync
    m_oWebView.ExecuteScript StrPtr(pvJsBuildCall(FuncName, vArgs)), oHandler
    '--- registered after ExecuteScript so a raised error cannot leak the token
    m_cScriptResults.Add True, "A" & jsRunAsync
    Exit Function
EH:
    PrintError "jsRunAsync"
End Function

Public Sub Shutdown()
    On Error Resume Next
    m_bShuttingDown = True
    If Not m_oWebView Is Nothing Then
        m_oWebView.RemoveHostObjectFromScript StrPtr("vbHost")
        Set m_oWebView = Nothing
    End If
    Set m_oWebView4 = Nothing
    Set m_oWebView11 = Nothing
    Set m_oWebView13 = Nothing
    If Not m_oController Is Nothing Then
        m_oController.Close
        Set m_oController = Nothing
    End If
    Set m_oEnvironment = Nothing
    If Not m_oEventSink Is Nothing Then
        m_oEventSink.Terminate
        Set m_oEventSink = Nothing
    End If
    If Not m_oFocusGotSink Is Nothing Then
        m_oFocusGotSink.Terminate
        Set m_oFocusGotSink = Nothing
    End If
    If Not m_oFocusLostSink Is Nothing Then
        m_oFocusLostSink.Terminate
        Set m_oFocusLostSink = Nothing
    End If
    Set m_oSettings = Nothing
    Set m_oSettings6 = Nothing
End Sub

'= private ===============================================================

Private Function pvCreateWebView2Environment( _
            sBrowserExecutableFolder As String, _
            ByVal sUserDataFolder As String, _
            oOptions As IVBWebView2EnvironmentOptions, _
            oHandler As IVBWebView2CreateEnvironmentCompletedHandler) As Long
    Const E_MOD_NOT_FOUND               As Long = &H8007007E
    Dim sLoader         As String

    On Error GoTo EH
    If GetModuleHandle(StrPtr("WebView2Loader.dll")) = 0 Then
        sLoader = LocateFile(App.Path & "\External\WebView2Loader.dll")
        '--- fail fast so a missing loader is not masked as BindTo timeout
        If LenB(sLoader) = 0 Then
            pvCreateWebView2Environment = E_MOD_NOT_FOUND
            Exit Function
        End If
        If LoadLibrary(StrPtr(sLoader)) = 0 Then
            pvCreateWebView2Environment = E_MOD_NOT_FOUND
            Exit Function
        End If
    End If
    '--- match RC6's default user data folder (%LOCALAPPDATA%\<EXEName>) instead
    '--- of the native default of "<exe>.WebView2" next to the executable
    If LenB(sUserDataFolder) = 0 Then
        sUserDataFolder = Environ$("LOCALAPPDATA") & "\" & UCase$(App.EXEName)
    End If
    pvCreateWebView2Environment = CreateCoreWebView2EnvironmentWithOptions(StrPtr(sBrowserExecutableFolder), StrPtr(sUserDataFolder), ObjPtr(oOptions), ObjPtr(oHandler))
    Exit Function
EH:
    pvCreateWebView2Environment = Err.Number
    PrintError "pvCreateWebView2Environment"
End Function

Private Function pvPumpUntil(bDone As Boolean, ByVal dblTimeout As Double) As Boolean
    Dim dblTimer        As Double

    On Error GoTo EH
    '--- 0 (or less) means fire-and-forget: caller wants to continue executing
    '--- asynchronously, so report success without pumping any messages
    If dblTimeout <= 0 Then
        pvPumpUntil = True
        Exit Function
    End If
    dblTimer = TimerEx + dblTimeout
    Do While TimerEx <= dblTimer
        If m_bShuttingDown Then
            Exit Do
        End If
        SpinThreadMessagePump Flush:=True
        If bDone Then
            pvPumpUntil = True
            Exit Do
        End If
    Loop
    Exit Function
EH:
    PrintError "pvPumpUntil"
End Function

Private Sub SpinThreadMessagePump(Optional ByVal Flush As Boolean, Optional ByVal hWnd As Long, Optional ByVal FromMsg As Long, Optional ByVal ToMsg As Long)
    Const PM_REMOVE     As Long = 1
    Const WM_NCMOUSEFIRST As Long = &HA0
    Const WM_NCMOUSELAST As Long = &HA9
    Const WM_KEYFIRST   As Long = &H100
    Const WM_KEYLAST    As Long = &H108
    Const WM_MOUSEFIRST As Long = &H200
    Const WM_MOUSELAST  As Long = &H209
    Dim uMsg            As APIMSG
    
    If Flush Then
        Call Sleep(1)
    End If
    Call CoWaitForMultipleHandles(0, 0, 0, 0, 0)
    Do While PeekMessage(VarPtr(uMsg), hWnd, FromMsg, ToMsg, PM_REMOVE) <> 0
        Call TranslateMessage(VarPtr(uMsg))
        If Flush And (uMsg.lMessage >= WM_KEYFIRST And uMsg.lMessage <= WM_KEYLAST _
                        Or uMsg.lMessage >= WM_MOUSEFIRST And uMsg.lMessage <= WM_MOUSELAST _
                        Or uMsg.lMessage >= WM_NCMOUSEFIRST And uMsg.lMessage <= WM_NCMOUSELAST) Then
            '--- flush
        Else
            Call DispatchMessage(VarPtr(uMsg))
        End If
        Call CoWaitForMultipleHandles(0, 0, 0, 0, 0)
    Loop
End Sub

Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Private Function pvToString(ByVal lPtr As Long) As String
    Call SysReAllocString(VarPtr(pvToString), lPtr)
    Call CoTaskMemFree(lPtr)
End Function

Private Function pvNextToken() As Currency
    m_cTokenSeq = m_cTokenSeq + 1
    pvNextToken = m_cTokenSeq
End Function

Private Function pvGetProfile7() As IVBCoreWebView2Profile7
    '--- the Set QIs the raw profile pointer for Profile7's IID and fails
    '--- soft on runtimes without browser extension support
    On Error Resume Next
    If Not m_oWebView13 Is Nothing Then
        Set pvGetProfile7 = m_oWebView13.Profile
    End If
End Function

Private Function pvExecScriptSync(sScript As String, vResult As Variant) As Boolean
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant
    Dim sJson           As String
    Dim lIdx            As Long

    On Error GoTo EH
    m_sJsLastError = vbNullString
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    m_oWebView.ExecuteScript StrPtr(sScript), oHandler
    '--- jsCallTimeOutSeconds 0 (or less) means fire-and-forget: report success
    '--- and continue, the result (if any) is discarded on arrival
    If m_dbljsCallTimeOutSeconds <= 0 Then
        pvExecScriptSync = True
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, m_dbljsCallTimeOutSeconds) Then
        m_cScriptResults.Remove "P" & cToken
        m_sJsLastError = "Timeout waiting for script result"
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If vItem(0) < 0 Then
        m_sJsLastError = "Error &H" & Hex$(vItem(0)) & " executing script"
        Exit Function
    End If
    sJson = vItem(1)
    lIdx = 1
    pvJsonParseValue sJson, lIdx, vResult
    pvExecScriptSync = True
    Exit Function
EH:
    m_sJsLastError = Err.Description
    PrintError "pvExecScriptSync"
End Function

Private Function pvJsBuildCall(sFuncName As String, vArgs As Variant) As String
    Dim lIdx            As Long
    Dim sParams         As String

    For lIdx = LBound(vArgs) To UBound(vArgs)
        If LenB(sParams) <> 0 Then
            sParams = sParams & ","
        End If
        sParams = sParams & pvJsonEncode(vArgs(lIdx))
    Next
    pvJsBuildCall = sFuncName & "(" & sParams & ")"
End Function

Private Function pvPumpUntilKey(cCol As Collection, sKey As String, ByVal dblTimeout As Double) As Boolean
    Dim dblTimer        As Double

    On Error GoTo EH
    dblTimer = TimerEx + dblTimeout
    Do While TimerEx <= dblTimer
        If m_bShuttingDown Then
            Exit Do
        End If
        SpinThreadMessagePump Flush:=True
        If pvSearchCollection(cCol, sKey) Then
            pvPumpUntilKey = True
            Exit Do
        End If
    Loop
    Exit Function
EH:
    PrintError "pvPumpUntilKey"
End Function

Private Function pvSearchCollection(cCol As IVbCollection, sKey As String) As Boolean
    pvSearchCollection = (cCol.Item(sKey) >= 0)
End Function

Private Sub pvAssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Private Function pvJsonEncode(vValue As Variant) As String
    Dim vElem           As Variant
    Dim sResult         As String
    Dim cCol            As Collection
    Dim aKeys()         As String
    Dim lIdx            As Long

    If IsArray(vValue) Then
        For Each vElem In vValue
            If LenB(sResult) <> 0 Then
                sResult = sResult & ","
            End If
            sResult = sResult & pvJsonEncode(vElem)
        Next
        pvJsonEncode = "[" & sResult & "]"
    ElseIf IsObject(vValue) Then
        If vValue Is Nothing Then
            pvJsonEncode = "null"
        ElseIf TypeName(vValue) = "Collection" Then
            Set cCol = vValue
            aKeys = CollectionAllKeys(cCol)
            For lIdx = 1 To cCol.Count
                If LenB(sResult) <> 0 Then
                    sResult = sResult & ","
                End If
                sResult = sResult & pvJsonEncodeString(aKeys(lIdx - 1)) & ":" & pvJsonEncode(cCol.Item(lIdx))
            Next
            pvJsonEncode = "{" & sResult & "}"
        Else
            Err.Raise 13, , "Unsupported object type for JSON encoding: " & TypeName(vValue)
        End If
    Else
        Select Case VarType(vValue)
        Case vbEmpty, vbNull, vbError
            pvJsonEncode = "null"
        Case vbBoolean
            If vValue Then
                pvJsonEncode = "true"
            Else
                pvJsonEncode = "false"
            End If
        Case vbString
            pvJsonEncode = pvJsonEncodeString(CStr(vValue))
        Case vbDate
            pvJsonEncode = pvJsonEncodeString(Format$(vValue, "yyyy-mm-dd\Thh:nn:ss"))
        Case Else
            pvJsonEncode = Trim$(Str$(vValue))
        End Select
    End If
End Function

Private Function pvJsonEncodeString(sValue As String) As String
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim sResult         As String

    sResult = """"
    For lIdx = 1 To Len(sValue)
        lChar = AscW(Mid$(sValue, lIdx, 1))
        Select Case lChar
        Case 34                                     '--- quote
            sResult = sResult & "\"""
        Case 92                                     '--- backslash
            sResult = sResult & "\\"
        Case 10
            sResult = sResult & "\n"
        Case 13
            sResult = sResult & "\r"
        Case 9
            sResult = sResult & "\t"
        Case Is < 32
            sResult = sResult & "\u" & Right$("000" & Hex$(lChar), 4)
        Case Else
            sResult = sResult & ChrW$(lChar)
        End Select
    Next
    pvJsonEncodeString = sResult & """"
End Function

Private Sub pvJsonParseValue(sText As String, lIdx As Long, vResult As Variant)
    Dim cDict           As Collection
    Dim cItems          As Collection
    Dim sKey            As String
    Dim vItem           As Variant
    Dim lJdx            As Long
    Dim lStart          As Long

    Call VariantClear(vResult)
    pvJsonSkipSpace sText, lIdx
    Select Case Mid$(sText, lIdx, 1)
    Case "{"
        Set cDict = New Collection
        lIdx = lIdx + 1
        pvJsonSkipSpace sText, lIdx
        If Mid$(sText, lIdx, 1) = "}" Then
            lIdx = lIdx + 1
        Else
            Do
                pvJsonSkipSpace sText, lIdx
                sKey = pvJsonParseString(sText, lIdx)
                pvJsonSkipSpace sText, lIdx
                '--- skip ":"
                lIdx = lIdx + 1
                pvJsonParseValue sText, lIdx, vItem
                '--- Collection keys are case-insensitive so "a" vs "A" (or a
                '--- plain duplicate key) would raise 457 - make last one win
                If pvSearchCollection(cDict, sKey) Then
                    cDict.Remove sKey
                End If
                cDict.Add vItem, sKey
                pvJsonSkipSpace sText, lIdx
                If Mid$(sText, lIdx, 1) <> "," Then
                    Exit Do
                End If
                lIdx = lIdx + 1
            Loop
            '--- skip "}"
            lIdx = lIdx + 1
        End If
        Set vResult = cDict
    Case "["
        Set cItems = New Collection
        lIdx = lIdx + 1
        pvJsonSkipSpace sText, lIdx
        If Mid$(sText, lIdx, 1) = "]" Then
            lIdx = lIdx + 1
        Else
            Do
                pvJsonParseValue sText, lIdx, vItem
                cItems.Add vItem
                pvJsonSkipSpace sText, lIdx
                If Mid$(sText, lIdx, 1) <> "," Then
                    Exit Do
                End If
                lIdx = lIdx + 1
            Loop
            '--- skip "]"
            lIdx = lIdx + 1
        End If
        If cItems.Count = 0 Then
            vResult = Array()
        Else
            ReDim vResult(0 To cItems.Count - 1)
            For lJdx = 1 To cItems.Count
                pvAssignVariant vResult(lJdx - 1), cItems.Item(lJdx)
            Next
        End If
    Case """"
        vResult = pvJsonParseString(sText, lIdx)
    Case "t"
        vResult = True
        lIdx = lIdx + 4
    Case "f"
        vResult = False
        lIdx = lIdx + 5
    Case "n"
        vResult = Null
        lIdx = lIdx + 4
    Case Else
        lStart = lIdx
        Do While InStr("+-0123456789.eE", Mid$(sText, lIdx, 1)) > 0 And lIdx <= Len(sText)
            lIdx = lIdx + 1
        Loop
        vResult = Val(Mid$(sText, lStart, lIdx - lStart))
    End Select
End Sub

Private Function pvJsonParseString(sText As String, lIdx As Long) As String
    Dim sChar           As String

    '--- skip opening quote
    lIdx = lIdx + 1
    Do While lIdx <= Len(sText)
        sChar = Mid$(sText, lIdx, 1)
        Select Case sChar
        Case """"
            lIdx = lIdx + 1
            Exit Do
        Case "\"
            lIdx = lIdx + 1
            sChar = Mid$(sText, lIdx, 1)
            Select Case sChar
            Case "n"
                pvJsonParseString = pvJsonParseString & vbLf
            Case "r"
                pvJsonParseString = pvJsonParseString & vbCr
            Case "t"
                pvJsonParseString = pvJsonParseString & vbTab
            Case "b"
                pvJsonParseString = pvJsonParseString & Chr$(8)
            Case "f"
                pvJsonParseString = pvJsonParseString & Chr$(12)
            Case "u"
                pvJsonParseString = pvJsonParseString & ChrW$(CLng("&H" & Mid$(sText, lIdx + 1, 4)))
                lIdx = lIdx + 4
            Case Else
                pvJsonParseString = pvJsonParseString & sChar
            End Select
            lIdx = lIdx + 1
        Case Else
            pvJsonParseString = pvJsonParseString & sChar
            lIdx = lIdx + 1
        End Select
    Loop
End Function

Private Sub pvJsonSkipSpace(sText As String, lIdx As Long)
    Do While lIdx <= Len(sText)
        Select Case Mid$(sText, lIdx, 1)
        Case " ", vbTab, vbCr, vbLf
            lIdx = lIdx + 1
        Case Else
            Exit Do
        End Select
    Loop
End Sub

Private Function CollectionAllKeys(oCol As Collection, Optional ByVal StartIndex As Long) As String()
    Const o_KeyPtr              As Long = &H10
    Const o_pNextIndexedItem    As Long = &H18
    Dim lPtr            As Long
    Dim aRetVal()       As String
    Dim lIdx            As Long
    Dim sTemp           As String
    
    If oCol.Count = 0 Then
        aRetVal = Split(vbNullString)
    Else
        ReDim aRetVal(StartIndex To StartIndex + oCol.Count - 1) As String
        lPtr = ObjPtr(oCol)
        For lIdx = LBound(aRetVal) To UBound(aRetVal)
            Call CopyMemory(lPtr, ByVal lPtr + o_pNextIndexedItem, 4)
            Call CopyMemory(ByVal VarPtr(sTemp), ByVal lPtr + o_KeyPtr, 4)
            aRetVal(lIdx) = sTemp
        Next
        Call CopyMemory(ByVal VarPtr(sTemp), 0&, 4)
    End If
    CollectionAllKeys = aRetVal
End Function

Private Function LocateFile(sFile As String) As String
    Dim sDir            As String
    Dim sName           As String
    Dim lPos            As Long
    
    If InStrRev(sFile, "\") > 0 Then
        sDir = Left$(sFile, InStrRev(sFile, "\"))
        sName = Mid$(sFile, InStrRev(sFile, "\") + 1)
        Do While Not FileExists(sDir & sName)
            If Len(sDir) > 1 Then
                lPos = InStrRev(sDir, "\", Len(sDir) - 1)
                If lPos > 0 Then
                    sDir = Left$(sDir, lPos)
                    If Left$(sDir, 2) = "\\" And InStrRev(sDir, "\", Len(sDir) - 1) <= 2 Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Loop
        LocateFile = sDir & sName
    ElseIf FileExists(sFile) Then
        LocateFile = sFile
    End If
End Function

Private Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(StrPtr(sFile)) = -1 Then ' INVALID_FILE_ATTRIBUTES
        FileExists = (Err.LastDllError = 32) ' ERROR_SHARING_VIOLATION
    Else
        FileExists = True
    End If
End Function

'= friend ================================================================

Friend Sub OnEnvironmentCreated(ByVal hResult As Long, ByVal oEnv As IVBCoreWebView2Environment)
    Dim oHandler        As cWebView2Callback

    On Error GoTo EH
    If hResult <> 0 Then
        m_hInitError = hResult
        m_bInitComplete = True
        Exit Sub
    End If
    Set m_oEnvironment = oEnv
    Set oHandler = New cWebView2Callback
    oHandler.Init Me
    m_oEnvironment.CreateCoreWebView2Controller m_hWndHost, oHandler
    Exit Sub
EH:
    m_hInitError = Err.Number
    m_bInitComplete = True
    PrintError "OnEnvironmentCreated"
End Sub

Friend Sub OnControllerCreated(ByVal hResult As Long, ByVal oCtrl As IVBCoreWebView2Controller)
    On Error GoTo EH
    If hResult <> 0 Then
        m_hInitError = hResult
        m_bInitComplete = True
        Exit Sub
    End If
    Set m_oController = oCtrl
    Set m_oWebView = m_oController.CoreWebView2
    Set m_oSettings = m_oWebView.Settings
    '--- Settings6 and ICoreWebView2_4 may be unavailable on older runtimes
    On Error Resume Next
    Set m_oSettings6 = m_oSettings
    Set m_oWebView4 = m_oWebView
    Set m_oWebView11 = m_oWebView
    Set m_oWebView13 = m_oWebView
    On Error GoTo EH
    Set m_oEventSink = New cWebView2Callback
    m_oEventSink.Init Me
    m_oWebView.add_NavigationStarting m_oEventSink
    m_oWebView.add_NavigationCompleted m_oEventSink
    m_oWebView.add_SourceChanged m_oEventSink
    m_oWebView.add_ProcessFailed m_oEventSink
    m_oWebView.add_WebMessageReceived m_oEventSink
    m_oWebView.add_ScriptDialogOpening m_oEventSink
    m_oWebView.add_PermissionRequested m_oEventSink
    m_oWebView.add_NewWindowRequested m_oEventSink
    m_oWebView.add_WebResourceRequested m_oEventSink
    m_oWebView.add_DocumentTitleChanged m_oEventSink
    m_oWebView.add_ContentLoading m_oEventSink
    m_oWebView.add_HistoryChanged m_oEventSink
    m_oWebView.add_WindowCloseRequested m_oEventSink
    m_oController.add_AcceleratorKeyPressed m_oEventSink
    m_oController.add_ZoomFactorChanged m_oEventSink
    m_oController.add_MoveFocusRequested m_oEventSink
    '--- GotFocus/LostFocus share one handler interface with no distinguishing
    '--- args, so each subscription gets its own discriminated sink instance
    Set m_oFocusGotSink = New cWebView2Callback
    m_oFocusGotSink.Init Me, 1
    m_oController.add_GotFocus m_oFocusGotSink
    Set m_oFocusLostSink = New cWebView2Callback
    m_oFocusLostSink.Init Me, 2
    m_oController.add_LostFocus m_oFocusLostSink
    If Not m_oWebView4 Is Nothing Then
        m_oWebView4.add_WebResourceResponseReceived m_oEventSink
        m_oWebView4.add_FrameCreated m_oEventSink
        m_oWebView4.add_DownloadStarting m_oEventSink
    End If
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.add_ContextMenuRequested m_oEventSink
    End If
    '--- RC6 compat: expose this instance to page script and inject the vbH()
    '--- accessor so code like vbH().RaiseMessageEvent(...) works unchanged
    AddObject "vbHost", m_oEventSink
    AddScriptToExecuteOnDocumentCreated "var vbH_; function vbH() { return vbH_ || (vbH_ = chrome.webview.hostObjects.sync.vbHost); }" & vbCrLf & _
        "window.addEventListener('contextmenu', function(e) { vbH().RaiseContextMenuEvent(e.screenX, e.screenY); });"
    SyncSizeToHostWindow
    m_oController.IsVisible = 1
    m_hInitError = 0
    m_bInitComplete = True
    RaiseEvent InitComplete
    Exit Sub
EH:
    m_hInitError = Err.Number
    m_bInitComplete = True
    PrintError "OnControllerCreated"
End Sub

Friend Sub SinkNavigationStarting(ByVal bIsUserInitiated As Boolean, ByVal bIsRedirected As Boolean, ByVal sUri As String, bCancel As Boolean)
    RaiseEvent NavigationStarting(bIsUserInitiated, bIsRedirected, sUri, bCancel)
End Sub

Friend Sub SinkNavigationCompleted(ByVal bIsSuccess As Boolean, ByVal lWebErrorStatus As Long)
    m_bNavDone = True
    m_bNavSuccess = bIsSuccess
    m_lNavErrorStatus = lWebErrorStatus
    RaiseEvent NavigationCompleted(bIsSuccess, lWebErrorStatus)
    If bIsSuccess Then
        RaiseEvent DocumentComplete
    End If
End Sub

Friend Sub SinkSourceChanged()
    RaiseEvent SourceChanged
End Sub

Friend Sub SinkProcessFailed()
    RaiseEvent ProcessFailed
End Sub

Friend Sub SinkScriptDialogOpening(ByVal eKind As eWebView2ScriptDialogKind, bAccept As Boolean, sResultText As String, sUri As String, sMessage As String, sDefaultText As String)
    RaiseEvent ScriptDialogOpening(eKind, bAccept, sResultText, sUri, sMessage, sDefaultText)
End Sub

Friend Sub SinkPermissionRequested(ByVal bIsUserInitiated As Boolean, eState As Long, sUri As String, ByVal lKind As Long)
    Dim eTemp           As eWebView2PermissionState

    '--- native kinds unknown to RC6 (>= 7) are reported as UNKNOWN_PERMISSION
    If lKind < PERMISSION_KIND_UNKNOWN_PERMISSION Or lKind > PERMISSION_KIND_CLIPBOARD_READ Then
        lKind = PERMISSION_KIND_UNKNOWN_PERMISSION
    End If
    eTemp = eState
    RaiseEvent PermissionRequested(bIsUserInitiated, eTemp, sUri, lKind)
    eState = eTemp
End Sub

Friend Sub SinkNewWindowRequested(ByVal bIsUserInitiated As Boolean, bHandled As Boolean, sUri As String, oFeatures As Collection)
    RaiseEvent NewWindowRequested(bIsUserInitiated, bHandled, sUri, oFeatures)
End Sub

Friend Sub SinkAcceleratorKeyPressed(ByVal eKeyState As eWebView2AccKeyState, ByVal bIsExtendedKey As Boolean, ByVal bWasKeyDown As Boolean, ByVal bIsKeyReleased As Boolean, ByVal bIsMenuKeyDown As Boolean, ByVal lRepeatCount As Long, ByVal lScanCode As Long, bIsHandled As Boolean)
    RaiseEvent AcceleratorKeyPressed(eKeyState, bIsExtendedKey, bWasKeyDown, bIsKeyReleased, bIsMenuKeyDown, lRepeatCount, lScanCode, bIsHandled)
End Sub

Friend Sub SinkWebResourceRequested(oArgs As IVBCoreWebView2WebResourceRequestedEventArgs)
    Dim oRequest        As IVBCoreWebView2WebResourceRequest
    Dim oResponse       As IVBCoreWebView2WebResourceResponse
    Dim ePermit         As eWebView2PermissionState

    On Error GoTo EH
    Set oRequest = oArgs.Request
    RaiseEvent WebResourceRequested(pvToString(oRequest.URI), oArgs.ResourceContext, ePermit)
    If ePermit = PERMISSION_STATE_DENY Then
        Set oResponse = m_oEnvironment.CreateWebResourceResponse(Nothing, 403, StrPtr("Forbidden"), 0)
        oArgs.Response = oResponse
    End If
    Exit Sub
EH:
    PrintError "SinkWebResourceRequested"
End Sub

Friend Sub SinkWebResourceResponseReceived(oArgs As IVBCoreWebView2WebResourceResponseReceivedEventArgs)
    Dim oRequest        As IVBCoreWebView2WebResourceRequest
    Dim oResponse       As IVBCoreWebView2WebResourceResponseView
    Dim oIter           As IVBCoreWebView2HttpHeadersCollectionIterator
    Dim cHeaders        As Collection
    Dim lName           As Long
    Dim lValue          As Long
    Dim sKey            As String
    Dim oTemp           As Collection
    
    On Error GoTo EH
    Set oRequest = oArgs.Request
    Set oResponse = oArgs.Response
    Set cHeaders = New Collection
    Set oIter = oResponse.Headers.GetIterator
    Do While oIter.HasCurrentHeader <> 0
        oIter.GetCurrentHeader lName, lValue
        sKey = pvToString(lName)
        If pvSearchCollection(cHeaders, sKey) Then
            If IsObject(cHeaders.Item(sKey)) Then
                Set oTemp = cHeaders.Item(sKey)
            Else
                Set oTemp = New Collection
                oTemp.Add cHeaders.Item(sKey)
                cHeaders.Remove sKey
                cHeaders.Add oTemp, sKey
            End If
            oTemp.Add pvToString(lValue)
        Else
            cHeaders.Add pvToString(lValue), sKey
        End If
        If oIter.MoveNext = 0 Then
            Exit Do
        End If
    Loop
    RaiseEvent WebResourceResponseReceived(pvToString(oRequest.URI), pvToString(oRequest.Method), oResponse.StatusCode, pvToString(oResponse.ReasonPhrase), cHeaders)
    Exit Sub
EH:
    PrintError "SinkWebResourceResponseReceived"
End Sub

Friend Sub SinkFrameCreated(sName As String)
    RaiseEvent FrameCreated(sName)
End Sub

Friend Sub SinkDocumentTitleChanged()
    RaiseEvent TitleChange(pvToString(m_oWebView.DocumentTitle))
End Sub

Friend Sub SinkContentLoading(ByVal bIsErrorPage As Boolean)
    RaiseEvent ContentLoading(bIsErrorPage)
End Sub

Friend Sub SinkHistoryChanged()
    RaiseEvent HistoryChanged
End Sub

Friend Sub SinkWindowCloseRequested()
    RaiseEvent WindowCloseRequested
End Sub

Friend Sub SinkZoomFactorChanged()
    RaiseEvent ZoomFactorChanged
End Sub

Friend Sub SinkMoveFocusRequested(ByVal lReason As Long, bHandled As Boolean)
    RaiseEvent MoveFocusRequested(lReason, bHandled)
End Sub

Friend Sub SinkContextMenuRequested(sPageURI As String, sLinkURI As String, sSelectionText As String, ByVal lScreenX As Long, ByVal lScreenY As Long, bHandled As Boolean)
    RaiseEvent ContextMenuRequested(sPageURI, sLinkURI, sSelectionText, lScreenX, lScreenY, bHandled)
End Sub

Friend Sub SinkDownloadStarting(oArgs As IVBCoreWebView2DownloadStartingEventArgs)
    Dim oOper           As IVBCoreWebView2DownloadOperation
    Dim sPath           As String
    Dim bCancelDialog   As Boolean
    Dim bCancelOp       As Boolean
    Dim sNewPath        As String

    On Error GoTo EH
    Set oOper = oArgs.DownloadOperation
    sPath = pvToString(oArgs.ResultFilePath)
    bCancelDialog = oArgs.Handled
    bCancelOp = oArgs.Cancel
    RaiseEvent DownloadStarting(pvToString(oOper.URI), pvToString(oOper.MimeType), sPath, Mid$(sPath, InStrRev(sPath, "\") + 1), bCancelDialog, bCancelOp, sNewPath)
    oArgs.Handled = -bCancelDialog
    oArgs.Cancel = -bCancelOp
    If LenB(sNewPath) <> 0 Then
        oArgs.ResultFilePath = StrPtr(sNewPath)
    End If
    Exit Sub
EH:
    PrintError "SinkDownloadStarting"
End Sub

Friend Sub SinkFocusChanged(ByVal bGotFocus As Boolean)
    If bGotFocus Then
        RaiseEvent GotFocus(m_eLastFocusReason)
        m_eLastFocusReason = FocusReason_PROGRAMMATIC
    Else
        RaiseEvent LostFocus(FocusReason_PROGRAMMATIC)
    End If
End Sub

Friend Sub SinkWebMessageReceived(ByVal bIsString As Boolean, sMsg As String, sJson As String)
    Dim oContent        As Collection
    Dim vResult         As Variant

    On Error GoTo EH
    If bIsString Then
        RaiseEvent JSMessage(sMsg, vbNullString, oContent)
    Else
        pvJsonParseValue sJson, 1, vResult
        If IsObject(vResult) Then
            Set oContent = vResult
        End If
        RaiseEvent JSMessage(vbNullString, sJson, oContent)
    End If
    Exit Sub
EH:
    PrintError "SinkWebMessageReceived"
End Sub

Friend Sub OnExecuteScriptCompleted(ByVal hResult As Long, sResultJson As String, ByVal cToken As Currency)
    Dim vResult         As Variant
    Dim sError          As String

    On Error GoTo EH
    If cToken = 0 Then
        '--- fire-and-forget ExecuteScript
    ElseIf pvSearchCollection(m_cScriptResults, "A" & cToken) Then
        m_cScriptResults.Remove "A" & cToken
        If hResult < 0 Then
            sError = "Error &H" & Hex$(hResult) & " executing script"
            m_sJsLastError = sError
        Else
            pvJsonParseValue sResultJson, 1, vResult
        End If
        RaiseEvent JSAsyncResult(vResult, cToken, sError)
    ElseIf pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        m_cScriptResults.Remove "P" & cToken
        m_cScriptResults.Add Array(hResult, sResultJson), "T" & cToken
    Else
        '--- unknown token (timed-out or fire-and-forget call) - discard result
    End If
    Exit Sub
EH:
    PrintError "OnExecuteScriptCompleted"
End Sub

Friend Sub OnGetBrowserExtensions(ByVal hResult As Long, oList As IVBCoreWebView2BrowserExtensionList, ByVal cToken As Currency)
    Dim cResult         As Collection
    Dim cItem           As Collection
    Dim oExt            As IVBCoreWebView2BrowserExtension
    Dim lIdx            As Long

    On Error GoTo EH
    If Not pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        '--- unknown token (timed-out or fire-and-forget call) - discard result
        Exit Sub
    End If
    m_cScriptResults.Remove "P" & cToken
    Set cResult = New Collection
    If hResult >= 0 And Not oList Is Nothing Then
        For lIdx = 0 To oList.Count - 1
            Set oExt = oList.GetValueAtIndex(lIdx)
            Set cItem = New Collection
            cItem.Add pvToString(oExt.Id), "Id"
            cItem.Add pvToString(oExt.Name), "Name"
            cItem.Add CBool(oExt.IsEnabled), "IsEnabled"
            cResult.Add cItem, cItem.Item("Id")
        Next
    End If
    m_cScriptResults.Add Array(hResult, cResult), "T" & cToken
    Exit Sub
EH:
    PrintError "OnGetBrowserExtensions"
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_dbljsCallTimeOutSeconds = 8
    Set m_cScriptResults = New Collection
End Sub

Private Sub Class_Terminate()
    Shutdown
End Sub
