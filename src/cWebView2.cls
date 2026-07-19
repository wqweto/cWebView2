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

Public Enum eWebView2HostResourceAccessKind
    HostResourceAccess_DENY = 0
    HostResourceAccess_ALLOW = 1
    HostResourceAccess_DENY_CORS = 2
End Enum

Public Enum eWebView2PreferredColorScheme
    ColorScheme_AUTO = 0
    ColorScheme_LIGHT = 1
    ColorScheme_DARK = 2
End Enum

Public Enum eWebView2TrackingPreventionLevel
    TrackingPrevention_NONE = 0
    TrackingPrevention_BASIC = 1
    TrackingPrevention_BALANCED = 2
    TrackingPrevention_STRICT = 3
End Enum

Public Enum eWebView2DownloadDialogCornerAlignment
    DownloadDialogCorner_TOP_LEFT = 0
    DownloadDialogCorner_TOP_RIGHT = 1
    DownloadDialogCorner_BOTTOM_LEFT = 2
    DownloadDialogCorner_BOTTOM_RIGHT = 3
End Enum

Public Enum eWebView2CookieSameSiteKind
    SameSite_NONE = 0
    SameSite_LAX = 1
    SameSite_STRICT = 2
End Enum

Public Enum eWebView2PrintDialogKind
    PrintDialog_BROWSER = 0
    PrintDialog_SYSTEM = 1
End Enum

Public Enum eWebView2BrowsingDataKinds
    BrowsingData_FILE_SYSTEMS = &H1
    BrowsingData_INDEXED_DB = &H2
    BrowsingData_LOCAL_STORAGE = &H4
    BrowsingData_WEB_SQL = &H8
    BrowsingData_CACHE_STORAGE = &H10
    BrowsingData_ALL_DOM_STORAGE = &H20
    BrowsingData_COOKIES = &H40
    BrowsingData_ALL_SITE = &H80
    BrowsingData_DISK_CACHE = &H100
    BrowsingData_DOWNLOAD_HISTORY = &H200
    BrowsingData_GENERAL_AUTOFILL = &H400
    BrowsingData_PASSWORD_AUTOSAVE = &H800
    BrowsingData_BROWSING_HISTORY = &H1000
    BrowsingData_SETTINGS = &H2000
    BrowsingData_ALL_PROFILE = &H3FFF
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
Public Event TitleChanged(ByVal Text As String)
Public Event ContentLoading(ByVal IsErrorPage As Boolean)
Public Event HistoryChanged()
Public Event WindowCloseRequested()
Public Event ZoomFactorChanged()
Public Event MoveFocusRequested(ByVal Reason As eWebView2FocusReason, Handled As Boolean)
Public Event ContextMenuRequested(ByVal PageURI As String, ByVal LinkURI As String, ByVal SelectionText As String, ByVal ScreenX As Long, ByVal ScreenY As Long, Handled As Boolean)
Public Event DOMContentLoaded()
Public Event StatusBarTextChanged(ByVal Text As String)
Public Event ContainsFullScreenElementChanged()
Public Event BasicAuthenticationRequested(ByVal URI As String, ByVal Challenge As String, UserName As String, Password As String, Cancel As Boolean)

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As APIRECT) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, ByVal lpsz As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader" (ByVal browserExecutableFolder As Long, ByVal userDataFolder As Long, ByVal environmentOptions As Long, ByVal environmentCreatedHandler As Long) As Long
Private Declare Function GetAvailableCoreWebView2BrowserVersionString Lib "WebView2Loader" (ByVal browserExecutableFolder As Long, versionInfo As Long) As Long
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
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventW" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageW" (ByVal lpString As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type APIMSG
    hWnd                As Long
    lMessage            As Long
    wParam              As Long
    lParam              As Long
    time                As Long
    pt                  As APIPOINT
End Type

Private Type APIDISPPARAMS
    rgvarg              As Long
    rgdispidNamedArgs   As Long
    cArgs               As Long
    cNamedArgs          As Long
End Type

'=========================================================================
' MST
'=========================================================================

Private Const SIGN_BIT                      As Long = &H80000000
Private Const PTR_SIZE                      As Long = 4

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    
'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oEnvironment              As IVBCoreWebView2Environment
Private m_oController               As IVBCoreWebView2Controller
Private m_oWebView                  As IVBCoreWebView2
Private m_oWebView4                 As IVBCoreWebView2_4
Private m_oWebView11                As IVBCoreWebView2_11
Private m_oWebView13                As IVBCoreWebView2_13
Private m_oWebView16                As IVBCoreWebView2_16
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
Private m_hPumpEvent                As Long
'--- per-document JS function table ({run, cbn, propGet, propLet}) received
'--- via SetFuncObj - the RC6-style synchronous bridge into the renderer
Private m_oFuncObj                  As Object
Private m_pTimer                    As IUnknown
Private m_cEventQueue               As Collection
Private m_lDeferMsg                 As Long

Private Enum eDeferredEvent
    ev_InitComplete = 1
    ev_DocumentComplete
    ev_NavigationCompleted
    ev_SourceChanged
    ev_ProcessFailed
    ev_JSMessage
    ev_JSAsyncResult
    ev_UserContextMenu
    ev_GotFocus
    ev_LostFocus
    ev_WebResourceResponseReceived
    ev_FrameCreated
    ev_TitleChanged
    ev_ContentLoading
    ev_HistoryChanged
    ev_WindowCloseRequested
    ev_ZoomFactorChanged
    ev_DOMContentLoaded
    ev_StatusBarTextChanged
    ev_ContainsFullScreenElementChanged
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & sFunction & "]"
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

Public Property Let HosthWnd(ByVal hValue As Long)
    m_hWndHost = hValue
    If Not m_oController Is Nothing Then
        m_oController.ParentWindow = hValue
    End If
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
    '--- dotted paths ride the bridge so the value can come back as a live JS
    '--- object; arbitrary expressions keep the ExecuteScript evaluation path
    If pvIsDottedPath(PropName) Then
        If pvBridgeInvoke("propGet", Array(PropName), Array(), jsProp) Then
            Exit Property
        End If
    End If
    pvExecScriptSync PropName, jsProp
End Property

Public Property Let jsProp(ByVal PropName As String, ByVal RHS As Variant)
    Dim vResult         As Variant

    If pvIsDottedPath(PropName) Then
        If pvBridgeInvoke("propLet", Array(PropName, RHS), Array(), vResult) Then
            Exit Property
        End If
    End If
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

Public Property Get IsBuiltInErrorPageEnabled() As Boolean
    If Not m_oSettings Is Nothing Then
        IsBuiltInErrorPageEnabled = m_oSettings.IsBuiltInErrorPageEnabled
    End If
End Property

Public Property Let IsBuiltInErrorPageEnabled(ByVal bValue As Boolean)
    If Not m_oSettings Is Nothing Then
        m_oSettings.IsBuiltInErrorPageEnabled = -bValue
    End If
End Property

Public Property Get ContainsFullScreenElement() As Boolean
    If Not m_oWebView Is Nothing Then
        ContainsFullScreenElement = m_oWebView.ContainsFullScreenElement
    End If
End Property

Public Property Get IsMuted() As Boolean
    If Not m_oWebView11 Is Nothing Then
        IsMuted = m_oWebView11.IsMuted
    End If
End Property

Public Property Let IsMuted(ByVal bValue As Boolean)
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.IsMuted = -bValue
    End If
End Property

Public Property Get IsDocumentPlayingAudio() As Boolean
    If Not m_oWebView11 Is Nothing Then
        IsDocumentPlayingAudio = m_oWebView11.IsDocumentPlayingAudio
    End If
End Property

Public Property Get IsDefaultDownloadDialogOpen() As Boolean
    If Not m_oWebView11 Is Nothing Then
        IsDefaultDownloadDialogOpen = m_oWebView11.IsDefaultDownloadDialogOpen
    End If
End Property

Public Property Get DefaultDownloadDialogCornerAlignment() As eWebView2DownloadDialogCornerAlignment
    If Not m_oWebView11 Is Nothing Then
        DefaultDownloadDialogCornerAlignment = m_oWebView11.DefaultDownloadDialogCornerAlignment
    End If
End Property

Public Property Let DefaultDownloadDialogCornerAlignment(ByVal eValue As eWebView2DownloadDialogCornerAlignment)
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.DefaultDownloadDialogCornerAlignment = eValue
    End If
End Property

Public Property Get DefaultDownloadDialogMargin() As APIPOINT
    If Not m_oWebView11 Is Nothing Then
        DefaultDownloadDialogMargin = m_oWebView11.DefaultDownloadDialogMargin
    End If
End Property

Public Property Let DefaultDownloadDialogMargin(uValue As APIPOINT)
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.SetDefaultDownloadDialogMargin uValue.x, uValue.y
    End If
End Property

Public Property Get StatusBarText() As String
    If Not m_oWebView13 Is Nothing Then
        StatusBarText = pvToString(m_oWebView13.StatusBarText)
    End If
End Property

Public Property Get BrowserVersion() As String
    If Not m_oEnvironment Is Nothing Then
        BrowserVersion = pvToString(m_oEnvironment.BrowserVersionString)
    End If
End Property

Public Property Get ProfileName() As String
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        ProfileName = pvToString(oProfile.ProfileName)
    End If
End Property

Public Property Get ProfilePath() As String
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        ProfilePath = pvToString(oProfile.ProfilePath)
    End If
End Property

Public Property Get IsInPrivateModeEnabled() As Boolean
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        IsInPrivateModeEnabled = oProfile.IsInPrivateModeEnabled
    End If
End Property

Public Property Get DefaultDownloadFolderPath() As String
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        DefaultDownloadFolderPath = pvToString(oProfile.DefaultDownloadFolderPath)
    End If
End Property

Public Property Let DefaultDownloadFolderPath(ByVal sValue As String)
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        oProfile.DefaultDownloadFolderPath = StrPtr(sValue)
    End If
End Property

Public Property Get PreferredColorScheme() As eWebView2PreferredColorScheme
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        PreferredColorScheme = oProfile.PreferredColorScheme
    End If
End Property

Public Property Let PreferredColorScheme(ByVal eValue As eWebView2PreferredColorScheme)
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        oProfile.PreferredColorScheme = eValue
    End If
End Property

Public Property Get PreferredTrackingPreventionLevel() As eWebView2TrackingPreventionLevel
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        PreferredTrackingPreventionLevel = oProfile.PreferredTrackingPreventionLevel
    End If
End Property

Public Property Let PreferredTrackingPreventionLevel(ByVal eValue As eWebView2TrackingPreventionLevel)
    Dim oProfile        As IVBCoreWebView2Profile7

    Set oProfile = pvGetProfile7()
    If Not oProfile Is Nothing Then
        oProfile.PreferredTrackingPreventionLevel = eValue
    End If
End Property

Private Property Get pvAddressOfTimerProc() As cWebView2
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
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
    '--- RC6 compat: 1 on success, 0 on any failure; underlying error codes
    '--- go to the debug console only
    hResult = pvCreateWebView2Environment(browserInstallPath, userDataFolder, oOptions, oHandler)
    If hResult <> 0 Then
        Debug.Print "Critical error: environment creation failed with &H" & Hex$(hResult) & " [BindTo]"
        Exit Function
    End If
    If Not pvPumpUntil(m_bInitComplete, SecondsToWaitForInitComplete) Then
        Debug.Print "Critical error: timeout waiting for init [BindTo]"
        Exit Function
    End If
    If m_hInitError <> 0 Then
        Debug.Print "Critical error: init failed with &H" & Hex$(m_hInitError) & " [BindTo]"
        Exit Function
    End If
    BindTo = 1
    Exit Function
EH:
    PrintError "BindTo"
End Function

Public Sub SyncSizeToHostWindow()
    Dim uRect           As APIRECT

    On Error GoTo EH
    Call GetClientRect(m_hWndHost, uRect)
    If Not m_oController Is Nothing Then
        m_oController.SetBounds 0, 0, uRect.Right, uRect.Bottom
        '--- keep popup/dialog placement correct after host window moves
        m_oController.NotifyParentWindowPositionChanged
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
    m_oController.NotifyParentWindowPositionChanged
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
    m_oWebView.Navigate StrPtr(URI)
    If Not pvPumpUntil(m_bNavDone, SecondsToWaitForDocumentComplete) Then
        Navigate = Empty
        Exit Function
    End If
    '--- RC6 compat: 1 whenever navigation completed within the wait - even
    '--- for an unsuccessful navigation (error page) - and Empty on timeout
    Navigate = 1
    Exit Function
EH:
    Navigate = Empty
    PrintError "Navigate"
End Function

Public Function NavigateToString(ByVal sHTMLContent As String, Optional ByVal SecondsToWaitForDocumentComplete As Double = 8) As Variant
    On Error GoTo EH
    m_bNavDone = False
    m_oWebView.NavigateToString StrPtr(sHTMLContent)
    If Not pvPumpUntil(m_bNavDone, SecondsToWaitForDocumentComplete) Then
        Exit Function
    End If
    NavigateToString = 1
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
    Dim lPtr            As Long
    Dim sPath           As String

    On Error GoTo EH
    If pvEnsureLoader() Then
        If GetAvailableCoreWebView2BrowserVersionString(0, lPtr) >= 0 Then
            VersionString = pvToString(lPtr)
        End If
    End If
    If LenB(VersionString) <> 0 Then
        sPath = Environ$("ProgramFiles(x86)")
        If LenB(sPath) = 0 Then
            sPath = Environ$("ProgramFiles")
        End If
        sPath = sPath & "\Microsoft\" & EdgeChannel & "\Application\" & VersionString
        If FileExists(sPath) Then
            GetMostRecentInstallPath = sPath
        End If
    End If
    Exit Function
EH:
    PrintError "GetMostRecentInstallPath"
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
    m_oWebView4.NavigateWithWebResourceRequest oRequest
    If pvPumpUntil(m_bNavDone, SecondsToWaitForRequestComplete) Then
        NavigateWithWebResourceRequest = 1
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

Public Sub SetVirtualHostNameToFolderMapping(ByVal HostName As String, ByVal FolderPath As String, Optional ByVal AccessKind As eWebView2HostResourceAccessKind = HostResourceAccess_DENY)
    On Error GoTo EH
    If Not m_oWebView4 Is Nothing Then
        m_oWebView4.SetVirtualHostNameToFolderMapping StrPtr(HostName), StrPtr(FolderPath), AccessKind
    End If
    Exit Sub
EH:
    PrintError "SetVirtualHostNameToFolderMapping"
End Sub

Public Sub ClearVirtualHostNameToFolderMapping(ByVal HostName As String)
    On Error GoTo EH
    If Not m_oWebView4 Is Nothing Then
        m_oWebView4.ClearVirtualHostNameToFolderMapping StrPtr(HostName)
    End If
    Exit Sub
EH:
    PrintError "ClearVirtualHostNameToFolderMapping"
End Sub

Public Sub Stop_()
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        m_oWebView.Stop
    End If
    Exit Sub
EH:
    PrintError "Stop_"
End Sub

Public Sub PostWebMessageAsString(ByVal sMessage As String)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        m_oWebView.PostWebMessageAsString StrPtr(sMessage)
    End If
    Exit Sub
EH:
    PrintError "PostWebMessageAsString"
End Sub

Public Sub PostWebMessageAsJson(ByVal sJson As String)
    On Error GoTo EH
    If Not m_oWebView Is Nothing Then
        m_oWebView.PostWebMessageAsJson StrPtr(sJson)
    End If
    Exit Sub
EH:
    PrintError "PostWebMessageAsJson"
End Sub

Public Sub OpenTaskManagerWindow()
    On Error GoTo EH
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.OpenTaskManagerWindow
    End If
    Exit Sub
EH:
    PrintError "OpenTaskManagerWindow"
End Sub

Public Sub OpenDefaultDownloadDialog()
    On Error GoTo EH
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.OpenDefaultDownloadDialog
    End If
    Exit Sub
EH:
    PrintError "OpenDefaultDownloadDialog"
End Sub

Public Sub CloseDefaultDownloadDialog()
    On Error GoTo EH
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.CloseDefaultDownloadDialog
    End If
    Exit Sub
EH:
    PrintError "CloseDefaultDownloadDialog"
End Sub

Public Sub SetBoundsAndZoomFactor(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal ZoomFactor As Double)
    On Error GoTo EH
    If Not m_oController Is Nothing Then
        m_oController.SetBoundsAndZoomFactor Left, Top, Right, Bottom, ZoomFactor
    End If
    Exit Sub
EH:
    PrintError "SetBoundsAndZoomFactor"
End Sub

Public Sub NotifyParentWindowPositionChanged()
    On Error GoTo EH
    If Not m_oController Is Nothing Then
        m_oController.NotifyParentWindowPositionChanged
    End If
    Exit Sub
EH:
    PrintError "NotifyParentWindowPositionChanged"
End Sub

Public Function TrySuspend(Optional ByVal SecondsToWaitForComplete As Double = 8) As Variant
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView4 Is Nothing Or m_oController Is Nothing Then
        Exit Function
    End If
    '--- the native call fails unless the webview is hidden first; visibility
    '--- is restored by ResumeFromSuspend (or automatically by the browser
    '--- when the webview becomes visible again)
    m_oController.IsVisible = 0
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    m_oWebView4.TrySuspend oHandler
    If SecondsToWaitForComplete <= 0 Then
        TrySuspend = True
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
    TrySuspend = (vItem(1) = "true")
    Exit Function
EH:
    PrintError "TrySuspend"
End Function

Public Sub ResumeFromSuspend()
    On Error GoTo EH
    If Not m_oWebView4 Is Nothing Then
        m_oWebView4.ResumeFromSuspend
    End If
    If Not m_oController Is Nothing Then
        m_oController.IsVisible = 1
    End If
    Exit Sub
EH:
    PrintError "ResumeFromSuspend"
End Sub

Public Property Get IsSuspended() As Boolean
    If Not m_oWebView4 Is Nothing Then
        IsSuspended = m_oWebView4.IsSuspended
    End If
End Property

Public Function ClearBrowsingData(ByVal DataKinds As eWebView2BrowsingDataKinds, Optional ByVal SecondsToWaitForComplete As Double = 8) As Variant
    ClearBrowsingData = pvClearBrowsingData(DataKinds, 0, 0, SecondsToWaitForComplete)
End Function

Public Function ClearBrowsingDataInTimeRange(ByVal DataKinds As eWebView2BrowsingDataKinds, ByVal StartTime As Date, ByVal EndTime As Date, Optional ByVal SecondsToWaitForComplete As Double = 8) As Variant
    Const SECS_PER_DAY                  As Long = 86400
    Dim dblStart        As Double
    Dim dblEnd          As Double

    '--- native times are seconds since the UNIX epoch
    dblStart = (StartTime - #1/1/1970#) * SECS_PER_DAY
    dblEnd = (EndTime - #1/1/1970#) * SECS_PER_DAY
    ClearBrowsingDataInTimeRange = pvClearBrowsingData(DataKinds, dblStart, dblEnd, SecondsToWaitForComplete, UseTimeRange:=True)
End Function

Public Function ClearBrowsingDataAll(Optional ByVal SecondsToWaitForComplete As Double = 8) As Variant
    ClearBrowsingDataAll = pvClearBrowsingData(0, 0, 0, SecondsToWaitForComplete, UseAll:=True)
End Function

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

    On Error GoTo EH
    sContent = sMsgContent
    '--- populate oJSONContent only when the content looks like a JSON object/array
    Select Case Left$(LTrim$(sContent), 1)
    Case "{", "["
        pvJsonParseValue sContent, 1, vResult
        If IsObject(vResult) Then
            Set oContent = vResult
        End If
    End Select
    '--- defer: this is called synchronously from page script over vbH(), so
    '--- raising JSMessage inline would deadlock a jsRun in the handler
    pvQueueEvent ev_JSMessage, CStr(sMsg), sContent, oContent
    Exit Sub
EH:
    PrintError "RaiseMessageEvent"
End Sub

Public Sub RaiseResultEvent(Res As Variant, ByVal Token As Variant, ByVal ErrMsg As Variant)
    Dim cToken          As Currency

    On Error GoTo EH
    cToken = CCur(Token)
    If pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        '--- nested completion of the pvBridgeInvoke call in flight below us
        '--- on the stack - stash the result (possibly a live JS object proxy)
        m_cScriptResults.Remove "P" & cToken
        m_cScriptResults.Add Array(0, Res, CStr(ErrMsg)), "T" & cToken
    Else
        pvQueueEvent ev_JSAsyncResult, Res, cToken, CStr(ErrMsg)
    End If
    Exit Sub
EH:
    PrintError "RaiseResultEvent"
End Sub

Public Sub SetFuncObj(FuncObj As Variant)
    On Error GoTo EH
    Set m_oFuncObj = FuncObj
    Exit Sub
EH:
    PrintError "SetFuncObj"
End Sub

Public Sub RaiseContextMenuEvent(ByVal ScreenX As Variant, ByVal ScreenY As Variant)
    On Error GoTo EH
    pvQueueEvent ev_UserContextMenu, CLng(ScreenX), CLng(ScreenY)
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
    pvJsonParseValue sJson, 1, CallDevToolsProtocolMethod
    Exit Function
EH:
    PrintError "CallDevToolsProtocolMethod"
End Function

Public Function CallDevToolsProtocolMethodForSession(ByVal SessionId As String, ByVal MethodName As String, ByVal ParamsAsJSON As String) As Variant
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant
    Dim sJson           As String

    On Error GoTo EH
    If m_oWebView11 Is Nothing Then
        Exit Function
    End If
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    m_oWebView11.CallDevToolsProtocolMethodForSession StrPtr(SessionId), StrPtr(MethodName), StrPtr(ParamsAsJSON), oHandler
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
    pvJsonParseValue sJson, 1, CallDevToolsProtocolMethodForSession
    Exit Function
EH:
    PrintError "CallDevToolsProtocolMethodForSession"
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
    Dim oSettings       As IVBCoreWebView2PrintSettings
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView11 Is Nothing Then
        Exit Function
    End If
    Set oSettings = pvBuildPrintSettings(Landscape, ScaleFactor, PageWidth, PageHeight, MarginTop, MarginBottom, MarginLeft, MarginRight, ShouldPrintBackgrounds, ShouldPrintSelectionOnly, ShouldPrintHeaderAndFooter, HeaderTitle, FooterUri)
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

Public Function PrintToPdfStream( _
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
        Optional ByVal FooterUri As String) As Byte()
    Dim oSettings       As IVBCoreWebView2PrintSettings
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    If m_oWebView16 Is Nothing Then
        Exit Function
    End If
    Set oSettings = pvBuildPrintSettings(Landscape, ScaleFactor, PageWidth, PageHeight, MarginTop, MarginBottom, MarginLeft, MarginRight, ShouldPrintBackgrounds, ShouldPrintSelectionOnly, ShouldPrintHeaderAndFooter, HeaderTitle, FooterUri)
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    m_oWebView16.PrintToPdfStream oSettings, oHandler
    If SecondsToWaitForPrintComplete <= 0 Then
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
    PrintToPdfStream = vItem(1)
    Exit Function
EH:
    PrintError "PrintToPdfStream"
End Function

Public Sub ShowPrintUI(Optional ByVal DialogKind As eWebView2PrintDialogKind = PrintDialog_BROWSER)
    On Error GoTo EH
    If Not m_oWebView16 Is Nothing Then
        m_oWebView16.ShowPrintUI DialogKind
    End If
    Exit Sub
EH:
    PrintError "ShowPrintUI"
End Sub

Public Function GetCookies(Optional ByVal URI As String, Optional ByVal SecondsToWaitForComplete As Double = 8) As Collection
    Dim oManager        As IVBCoreWebView2CookieManager
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant

    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If oManager Is Nothing Then
        Exit Function
    End If
    cToken = pvNextToken()
    sKey = "T" & cToken
    Set oHandler = New cWebView2Callback
    oHandler.Init Me, cToken
    oManager.GetCookies StrPtr(URI), oHandler
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
    Set GetCookies = vItem(1)
    Exit Function
EH:
    PrintError "GetCookies"
End Function

Public Sub AddOrUpdateCookie( _
        ByVal Name As String, _
        ByVal Value As String, _
        ByVal Domain As String, _
        Optional ByVal Path As String = "/", _
        Optional ByVal Expires As Date, _
        Optional ByVal IsSecure As Boolean, _
        Optional ByVal IsHttpOnly As Boolean, _
        Optional ByVal SameSite As eWebView2CookieSameSiteKind = SameSite_LAX)
    Const SECS_PER_DAY                  As Long = 86400
    Dim oManager        As IVBCoreWebView2CookieManager
    Dim oCookie         As IVBCoreWebView2Cookie

    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If oManager Is Nothing Then
        Exit Sub
    End If
    Set oCookie = oManager.CreateCookie(StrPtr(Name), StrPtr(Value), StrPtr(Domain), StrPtr(Path))
    '--- zero date means session cookie (native default Expires of -1)
    If Expires <> 0 Then
        oCookie.Expires = (Expires - #1/1/1970#) * SECS_PER_DAY
    End If
    oCookie.IsSecure = -IsSecure
    oCookie.IsHttpOnly = -IsHttpOnly
    oCookie.SameSite = SameSite
    oManager.AddOrUpdateCookie oCookie
    Exit Sub
EH:
    PrintError "AddOrUpdateCookie"
End Sub

Public Sub DeleteCookies(ByVal Name As String, ByVal URI As String)
    Dim oManager        As IVBCoreWebView2CookieManager

    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If Not oManager Is Nothing Then
        '--- cookies are matched by name AND URI natively - an empty URI
        '--- deletes nothing, hence the parameter is not optional
        oManager.DeleteCookies StrPtr(Name), StrPtr(URI)
    End If
    Exit Sub
EH:
    PrintError "DeleteCookies"
End Sub

Public Sub DeleteAllCookies()
    Dim oManager        As IVBCoreWebView2CookieManager

    On Error GoTo EH
    Set oManager = pvGetCookieManager()
    If Not oManager Is Nothing Then
        oManager.DeleteAllCookies
    End If
    Exit Sub
EH:
    PrintError "DeleteAllCookies"
End Sub

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
    Dim vArgs           As Variant
    Dim hResult         As Long

    '--- raw IDispatch with LOCALE_USER_DEFAULT
    vArgs = P
    hResult = pvDispatchInvoke(Obj, ProcName, CallType, vArgs, CallByName)
    If hResult < 0 Then
        Err.Raise hResult
    End If
End Function

Public Function jsCallByName(ByVal Obj As Object, ByVal MethodName As String, ByVal CallType As VbCallType, ParamArray P() As Variant) As Variant
    Dim vArgs           As Variant

    '--- dispatches on a live JS object over the sync bridge; on failure the
    '--- result stays Empty and jsLastError holds the reason (there is no
    '--- ExecuteScript fallback - live JS proxies only travel over the bridge)
    vArgs = P
    pvBridgeInvoke "cbn", Array(Obj, CallType, MethodName), vArgs, jsCallByName
End Function

Public Function jsRun(ByVal FuncName As String, ParamArray P() As Variant) As Variant
    Dim vArgs           As Variant

    vArgs = P
    If Not pvBridgeInvoke("run", Array(FuncName), vArgs, jsRun) Then
        pvExecScriptSync pvJsBuildCall(FuncName, vArgs), jsRun
    End If
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
    If Not m_pTimer Is Nothing Then
        Call TerminateFireOnceTimerThunk(m_pTimer, Me)
        Set m_pTimer = Nothing
    End If
    Set m_cEventQueue = Nothing
    Set m_oFuncObj = Nothing
    If Not m_oWebView Is Nothing Then
        m_oWebView.RemoveHostObjectFromScript StrPtr("vbHost")
        Set m_oWebView = Nothing
    End If
    Set m_oWebView4 = Nothing
    Set m_oWebView11 = Nothing
    Set m_oWebView13 = Nothing
    Set m_oWebView16 = Nothing
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
    If m_hPumpEvent <> 0 Then
        Call CloseHandle(m_hPumpEvent)
        m_hPumpEvent = 0
    End If
End Sub

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Dim cBatch          As Collection
    Dim vEvent          As Variant
    Dim oContent        As Collection
    Dim vAsync          As Variant

    On Error GoTo EH
    '--- fire-once: release the timer up front so events queued while raising
    '--- (e.g. a jsRun that itself pumps) re-arm a fresh timer for the next tick
    Set m_pTimer = Nothing
    Set cBatch = m_cEventQueue
    Set m_cEventQueue = Nothing
    If cBatch Is Nothing Then
        Exit Function
    End If
    For Each vEvent In cBatch
        Select Case vEvent(0)
        Case ev_InitComplete
            RaiseEvent InitComplete
        Case ev_DocumentComplete
            RaiseEvent DocumentComplete
        Case ev_NavigationCompleted
            RaiseEvent NavigationCompleted(vEvent(1), vEvent(2))
        Case ev_SourceChanged
            RaiseEvent SourceChanged
        Case ev_ProcessFailed
            RaiseEvent ProcessFailed
        Case ev_JSMessage
            Set oContent = Nothing
            If IsObject(vEvent(3)) Then
                Set oContent = vEvent(3)
            End If
            RaiseEvent JSMessage(CStr(vEvent(1)), CStr(vEvent(2)), oContent)
        Case ev_JSAsyncResult
            pvAssignVariant vAsync, vEvent(1)
            RaiseEvent JSAsyncResult(vAsync, CCur(vEvent(2)), CStr(vEvent(3)))
        Case ev_UserContextMenu
            RaiseEvent UserContextMenu(vEvent(1), vEvent(2))
        Case ev_GotFocus
            RaiseEvent GotFocus(vEvent(1))
        Case ev_LostFocus
            RaiseEvent LostFocus(vEvent(1))
        Case ev_WebResourceResponseReceived
            Set oContent = Nothing
            If IsObject(vEvent(5)) Then
                Set oContent = vEvent(5)
            End If
            RaiseEvent WebResourceResponseReceived(CStr(vEvent(1)), CStr(vEvent(2)), vEvent(3), CStr(vEvent(4)), oContent)
        Case ev_FrameCreated
            RaiseEvent FrameCreated(CStr(vEvent(1)))
        Case ev_TitleChanged
            RaiseEvent TitleChanged(CStr(vEvent(1)))
        Case ev_ContentLoading
            RaiseEvent ContentLoading(vEvent(1))
        Case ev_HistoryChanged
            RaiseEvent HistoryChanged
        Case ev_WindowCloseRequested
            RaiseEvent WindowCloseRequested
        Case ev_ZoomFactorChanged
            RaiseEvent ZoomFactorChanged
        Case ev_DOMContentLoaded
            RaiseEvent DOMContentLoaded
        Case ev_StatusBarTextChanged
            RaiseEvent StatusBarTextChanged(CStr(vEvent(1)))
        Case ev_ContainsFullScreenElementChanged
            RaiseEvent ContainsFullScreenElementChanged
        End Select
    Next
    Exit Function
EH:
    PrintError "TimerProc"
End Function

'= private ===============================================================

Private Sub pvQueueEvent(ParamArray P() As Variant)
    Dim vArgs           As Variant

    vArgs = P
    If m_cEventQueue Is Nothing Then
        Set m_cEventQueue = New Collection
    End If
    m_cEventQueue.Add vArgs
    If m_hPumpEvent <> 0 Then
        Call PostMessage(m_hWndHost, m_lDeferMsg, 0, 0)
    ElseIf m_pTimer Is Nothing Then
        Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc())
    End If
End Sub

Private Function pvCreateWebView2Environment( _
            sBrowserExecutableFolder As String, _
            ByVal sUserDataFolder As String, _
            oOptions As IVBWebView2EnvironmentOptions, _
            oHandler As IVBWebView2CreateEnvironmentCompletedHandler) As Long
    Const E_MOD_NOT_FOUND               As Long = &H8007007E

    On Error GoTo EH
    '--- fail fast so a missing loader is not masked as BindTo timeout
    If Not pvEnsureLoader() Then
        pvCreateWebView2Environment = E_MOD_NOT_FOUND
        Exit Function
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

    If m_hPumpEvent = 0 Then
        m_hPumpEvent = CreateEvent(0, 1, 0, 0)
    End If
    If Flush Then
        Call Sleep(1)
    End If
    Call CoWaitForMultipleHandles(0, 0, 1, m_hPumpEvent, 0)
    Do While PeekMessage(VarPtr(uMsg), hWnd, FromMsg, ToMsg, PM_REMOVE) <> 0
        Call TranslateMessage(VarPtr(uMsg))
        If uMsg.lMessage = m_lDeferMsg Then
            TimerProc
        ElseIf Flush And (uMsg.lMessage >= WM_KEYFIRST And uMsg.lMessage <= WM_KEYLAST _
                        Or uMsg.lMessage >= WM_MOUSEFIRST And uMsg.lMessage <= WM_MOUSELAST _
                        Or uMsg.lMessage >= WM_NCMOUSEFIRST And uMsg.lMessage <= WM_NCMOUSELAST) Then
            '--- flush
        Else
            Call DispatchMessage(VarPtr(uMsg))
        End If
        Call CoWaitForMultipleHandles(0, 0, 1, m_hPumpEvent, 0)
    Loop
    If m_hPumpEvent <> 0 Then
        Call CloseHandle(m_hPumpEvent)
        m_hPumpEvent = 0
    End If
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

Private Function pvEnsureLoader() As Boolean
    Dim sLoader         As String

    If GetModuleHandle("WebView2Loader.dll") <> 0 Then
        pvEnsureLoader = True
        Exit Function
    End If
    sLoader = LocateFile(App.Path & "\External\WebView2Loader.dll")
    If LenB(sLoader) <> 0 Then
        pvEnsureLoader = (LoadLibrary(StrPtr(sLoader)) <> 0)
    End If
End Function

Private Function pvClearBrowsingData(ByVal lDataKinds As Long, ByVal dblStartTime As Double, ByVal dblEndTime As Double, ByVal dblTimeout As Double, Optional ByVal UseTimeRange As Boolean, Optional ByVal UseAll As Boolean) As Variant
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
    If UseAll Then
        oProfile.ClearBrowsingDataAll oHandler
    ElseIf UseTimeRange Then
        oProfile.ClearBrowsingDataInTimeRange lDataKinds, dblStartTime, dblEndTime, oHandler
    Else
        oProfile.ClearBrowsingData lDataKinds, oHandler
    End If
    If dblTimeout <= 0 Then
        pvClearBrowsingData = True
        Exit Function
    End If
    m_cScriptResults.Add True, "P" & cToken
    If Not pvPumpUntilKey(m_cScriptResults, sKey, dblTimeout) Then
        m_cScriptResults.Remove "P" & cToken
        Exit Function
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    pvClearBrowsingData = (vItem(0) >= 0)
    Exit Function
EH:
    PrintError "pvClearBrowsingData"
End Function

Private Function pvGetProfile7() As IVBCoreWebView2Profile7
    '--- the Set QIs the raw profile pointer for Profile7's IID and fails
    '--- soft on runtimes without browser extension support
    On Error Resume Next
    If Not m_oWebView13 Is Nothing Then
        Set pvGetProfile7 = m_oWebView13.Profile
    End If
End Function

Private Function pvDispatchInvoke(oDisp As IVBDispatch, sName As String, ByVal eCallType As VbCallType, vArgs As Variant, vResult As Variant) As Long
    Const LOCALE_USER_DEFAULT           As Long = &H400
    Const DISPID_PROPERTYPUT            As Long = -3
    Const DISPATCH_PROPERTYPUT          As Long = 4
    Const DISPATCH_PROPERTYPUTREF       As Long = 8
    Dim uIIDNull(0 To 3) As Long
    Dim uParams         As APIDISPPARAMS
    Dim aRev()          As Variant
    Dim lNamePtr        As Long
    Dim lDispId         As Long
    Dim lNamedPut       As Long
    Dim lFlags          As Long
    Dim lCount          As Long
    Dim lIdx            As Long

    '--- oDisp is the object QI'd to IDispatch by the caller's implicit cast;
    '--- GetIDsOfNames/Invoke with LOCALE_USER_DEFAULT - JS host-object proxies
    '--- reject 0 with DISP_E_UNKNOWNLCID. VbCallType maps 1:1 onto DISPATCH_*
    Call VariantClear(vResult)
    lNamePtr = StrPtr(sName)
    pvDispatchInvoke = oDisp.GetIDsOfNames(VarPtr(uIIDNull(0)), VarPtr(lNamePtr), 1, LOCALE_USER_DEFAULT, VarPtr(lDispId))
    If pvDispatchInvoke < 0 Then
        Exit Function
    End If
    If IsArray(vArgs) Then
        lCount = UBound(vArgs) - LBound(vArgs) + 1
    End If
    If lCount > 0 Then
        '--- DISPPARAMS wants the args in reverse order
        ReDim aRev(0 To lCount - 1)
        For lIdx = 0 To lCount - 1
            pvAssignVariant aRev(lCount - 1 - lIdx), vArgs(LBound(vArgs) + lIdx)
        Next
        uParams.rgvarg = VarPtr(aRev(0))
        uParams.cArgs = lCount
    End If
    lFlags = eCallType
    If lFlags = DISPATCH_PROPERTYPUT Or lFlags = DISPATCH_PROPERTYPUTREF Then
        lNamedPut = DISPID_PROPERTYPUT
        uParams.rgdispidNamedArgs = VarPtr(lNamedPut)
        uParams.cNamedArgs = 1
    End If
    pvDispatchInvoke = oDisp.Invoke(lDispId, VarPtr(uIIDNull(0)), LOCALE_USER_DEFAULT, lFlags, VarPtr(uParams), VarPtr(vResult), 0, 0)
End Function

Private Function pvBridgeInvoke(sEntry As String, vPre As Variant, vArgs As Variant, vResult As Variant) As Boolean
    Dim aLogical()      As Variant
    Dim vRetVal         As Variant
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim hResult         As Long

    If m_oFuncObj Is Nothing Then
        Exit Function
    End If
    On Error GoTo EH
    m_sJsLastError = vbNullString
    cToken = pvNextToken()
    sKey = "T" & cToken
    '--- logical call args: pre-args, token, trailing args - so the JS glue
    '--- sees (funcName, token, arg0, arg1, ...); pvDispatchInvoke reverses
    '--- them into DISPPARAMS order
    lCount = UBound(vPre) + UBound(vArgs) + 3
    ReDim aLogical(0 To lCount - 1)
    For lIdx = 0 To UBound(vPre)
        pvAssignVariant aLogical(lIdx), vPre(lIdx)
    Next
    aLogical(UBound(vPre) + 1) = CDbl(cToken)
    For lIdx = 0 To UBound(vArgs)
        pvAssignVariant aLogical(UBound(vPre) + 2 + lIdx), vArgs(lIdx)
    Next
    m_cScriptResults.Add True, "P" & cToken
    hResult = pvDispatchInvoke(m_oFuncObj, sEntry, VbMethod, aLogical, vRetVal)
    If hResult < 0 Then
        Err.Raise hResult
    End If
    If Not pvPumpUntilKey(m_cScriptResults, sKey, m_dbljsCallTimeOutSeconds) Then
        m_sJsLastError = "Timeout waiting for JS bridge result"
        GoTo QH
    End If
    vItem = m_cScriptResults.Item(sKey)
    m_cScriptResults.Remove sKey
    If LenB(vItem(2)) <> 0 Then
        m_sJsLastError = vItem(2)
    Else
        pvAssignVariant vResult, vItem(1)
    End If
    pvBridgeInvoke = True
QH:
    If pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        m_cScriptResults.Remove "P" & cToken
    End If
    Exit Function
EH:
    m_sJsLastError = "Error &H" & Hex$(Err.Number) & " invoking JS bridge"
    PrintError "pvBridgeInvoke"
    Resume QH
End Function

Private Function pvIsDottedPath(sPath As String) As Boolean
    Dim lIdx            As Long

    If LenB(sPath) = 0 Then
        Exit Function
    End If
    For lIdx = 1 To Len(sPath)
        Select Case AscW(Mid$(sPath, lIdx, 1))
        Case 48 To 57, 65 To 90, 97 To 122, 95, 36, 46 '--- 0-9 A-Z a-z _ $ .
        Case Else
            Exit Function
        End Select
    Next
    pvIsDottedPath = True
End Function

Private Function pvGetCookieManager() As IVBCoreWebView2CookieManager
    On Error Resume Next
    If Not m_oWebView4 Is Nothing Then
        Set pvGetCookieManager = m_oWebView4.CookieManager
    End If
End Function

Private Function pvBuildPrintSettings( _
            ByVal bLandscape As Boolean, _
            ByVal dblScaleFactor As Double, _
            ByVal dblPageWidth As Double, _
            ByVal dblPageHeight As Double, _
            ByVal dblMarginTop As Double, _
            ByVal dblMarginBottom As Double, _
            ByVal dblMarginLeft As Double, _
            ByVal dblMarginRight As Double, _
            ByVal bShouldPrintBackgrounds As Boolean, _
            ByVal bShouldPrintSelectionOnly As Boolean, _
            ByVal bShouldPrintHeaderAndFooter As Boolean, _
            sHeaderTitle As String, _
            sFooterUri As String) As IVBCoreWebView2PrintSettings
    Const ORIENT_LANDSCAPE              As Long = 1
    Dim oEnv6           As IVBCoreWebView2Environment6
    Dim oSettings       As IVBCoreWebView2PrintSettings

    On Error Resume Next
    Set oEnv6 = m_oEnvironment
    On Error GoTo EH
    If oEnv6 Is Nothing Then
        Exit Function
    End If
    Set oSettings = oEnv6.CreatePrintSettings()
    If bLandscape Then
        oSettings.Orientation = ORIENT_LANDSCAPE
    End If
    If dblScaleFactor > 0 Then
        oSettings.ScaleFactor = dblScaleFactor
    End If
    If dblPageWidth > 0 Then
        oSettings.PageWidth = dblPageWidth
    End If
    If dblPageHeight > 0 Then
        oSettings.PageHeight = dblPageHeight
    End If
    If dblMarginTop >= 0 Then
        oSettings.MarginTop = dblMarginTop
    End If
    If dblMarginBottom >= 0 Then
        oSettings.MarginBottom = dblMarginBottom
    End If
    If dblMarginLeft >= 0 Then
        oSettings.MarginLeft = dblMarginLeft
    End If
    If dblMarginRight >= 0 Then
        oSettings.MarginRight = dblMarginRight
    End If
    oSettings.ShouldPrintBackgrounds = -bShouldPrintBackgrounds
    oSettings.ShouldPrintSelectionOnly = -bShouldPrintSelectionOnly
    oSettings.ShouldPrintHeaderAndFooter = -bShouldPrintHeaderAndFooter
    If LenB(sHeaderTitle) <> 0 Then
        oSettings.HeaderTitle = StrPtr(sHeaderTitle)
    End If
    If LenB(sFooterUri) <> 0 Then
        oSettings.FooterUri = StrPtr(sFooterUri)
    End If
    Set pvBuildPrintSettings = oSettings
    Exit Function
EH:
    PrintError "pvBuildPrintSettings"
End Function

Private Function pvExecScriptSync(sScript As String, vResult As Variant) As Boolean
    Dim oHandler        As cWebView2Callback
    Dim cToken          As Currency
    Dim sKey            As String
    Dim vItem           As Variant
    Dim sJson           As String

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
    pvJsonParseValue sJson, 1, vResult
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
        '--- RC6 compat: integral JSON numbers travel as Long
        If vResult = Int(vResult) And Abs(vResult) <= 2147483647 Then
            vResult = CLng(vResult)
        End If
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
    Const NULL_PTR              As Long = 0
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
        Call CopyMemory(ByVal VarPtr(sTemp), NULL_PTR, 4)
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
    Set m_oWebView16 = m_oWebView
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
    m_oWebView.add_ContainsFullScreenElementChanged m_oEventSink
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
        m_oWebView4.add_DOMContentLoaded m_oEventSink
    End If
    If Not m_oWebView11 Is Nothing Then
        m_oWebView11.add_ContextMenuRequested m_oEventSink
        m_oWebView11.add_BasicAuthenticationRequested m_oEventSink
    End If
    If Not m_oWebView13 Is Nothing Then
        m_oWebView13.add_StatusBarTextChanged m_oEventSink
    End If
    '--- RC6 compat: expose this instance to page script and inject the vbH()
    '--- accessor so code like vbH().RaiseMessageEvent(...) works unchanged
    AddObject "vbHost", m_oEventSink
    AddScriptToExecuteOnDocumentCreated "var vbH_; function vbH() { return vbH_ || (vbH_ = chrome.webview.hostObjects.sync.vbHost); }" & vbCrLf & _
        "window.addEventListener('contextmenu', function(e) { vbH().RaiseContextMenuEvent(e.screenX, e.screenY); });"
    '--- RC6-style bridge: hand a JS function table to the host so jsRun/jsProp/
    '--- jsCallByName dispatch into the renderer and report back through
    '--- RaiseResultEvent, which returns LIVE JS object proxies (not the JSON
    '--- copies ExecuteScript yields) and is what makes jsCallByName possible
    AddScriptToExecuteOnDocumentCreated "(function() {" & vbCrLf & _
        "function _walk(path) { var ctx = window, ns = path.split('.'), name = ns.pop(); for (var i = 0; i < ns.length; i++) { ctx = ctx[ns[i]]; } return [ctx, name]; }" & vbCrLf & _
        "function _run(funcName, token) { var eMsg = '', res = null, args = Array.prototype.slice.call(arguments, 2);" & vbCrLf & _
        "  try { var w = _walk(funcName); res = w[0][w[1]].apply(w[0], args); } catch(e) { eMsg = 'Err in ' + funcName + '(): ' + e.message; }" & vbCrLf & _
        "  try { vbH().RaiseResultEvent(res, token, eMsg); } catch(e) {} }" & vbCrLf & _
        "function _cbn(obj, callType, methodName, token) { var eMsg = '', res = null;" & vbCrLf & _
        "  try { switch (callType) {" & vbCrLf & _
        "  case 1: res = obj[methodName].apply(obj, Array.prototype.slice.call(arguments, 4)); break;" & vbCrLf & _
        "  case 2: res = obj[methodName]; break;" & vbCrLf & _
        "  case 4: case 8: res = obj[methodName] = arguments[4]; break;" & vbCrLf & _
        "  } } catch(e) { eMsg = 'Err in ' + methodName + '(): ' + e.message; }" & vbCrLf & _
        "  try { vbH().RaiseResultEvent(res, token, eMsg); } catch(e) {} }" & vbCrLf & _
        "function _propGet(propName, token) { var eMsg = '', res = null;" & vbCrLf & _
        "  try { var w = _walk(propName); res = w[0][w[1]]; } catch(e) { eMsg = 'Err in ' + propName + '-Get: ' + e.message; }" & vbCrLf & _
        "  try { vbH().RaiseResultEvent(res, token, eMsg); } catch(e) {} }" & vbCrLf & _
        "function _propLet(propName, RHS, token) { var eMsg = '', res = true;" & vbCrLf & _
        "  try { var w = _walk(propName); w[0][w[1]] = RHS; } catch(e) { eMsg = 'Err in ' + propName + '-Let: ' + e.message; }" & vbCrLf & _
        "  try { vbH().RaiseResultEvent(res, token, eMsg); } catch(e) {} }" & vbCrLf & _
        "try { vbH().SetFuncObj({ run: _run, cbn: _cbn, propGet: _propGet, propLet: _propLet }); } catch(e) {}" & vbCrLf & _
        "})();"
    SyncSizeToHostWindow
    m_oController.IsVisible = 1
    m_hInitError = 0
    m_bInitComplete = True
    pvQueueEvent ev_InitComplete
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
    '--- the done-flag stays synchronous so a blocking Navigate returns here;
    '--- the events are deferred so they fire after Navigate returns (like RC6)
    m_bNavDone = True
    m_bNavSuccess = bIsSuccess
    m_lNavErrorStatus = lWebErrorStatus
    pvQueueEvent ev_NavigationCompleted, bIsSuccess, lWebErrorStatus
    If bIsSuccess Then
        pvQueueEvent ev_DocumentComplete
    End If
End Sub

Friend Sub SinkSourceChanged()
    pvQueueEvent ev_SourceChanged
End Sub

Friend Sub SinkProcessFailed()
    pvQueueEvent ev_ProcessFailed
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
    pvQueueEvent ev_WebResourceResponseReceived, pvToString(oRequest.URI), pvToString(oRequest.Method), oResponse.StatusCode, pvToString(oResponse.ReasonPhrase), cHeaders
    Exit Sub
EH:
    PrintError "SinkWebResourceResponseReceived"
End Sub

Friend Sub SinkFrameCreated(sName As String)
    pvQueueEvent ev_FrameCreated, sName
End Sub

Friend Sub SinkDocumentTitleChanged()
    '--- read the title synchronously here; the event is deferred
    pvQueueEvent ev_TitleChanged, pvToString(m_oWebView.DocumentTitle)
End Sub

Friend Sub SinkContentLoading(ByVal bIsErrorPage As Boolean)
    pvQueueEvent ev_ContentLoading, bIsErrorPage
End Sub

Friend Sub SinkHistoryChanged()
    pvQueueEvent ev_HistoryChanged
End Sub

Friend Sub SinkWindowCloseRequested()
    pvQueueEvent ev_WindowCloseRequested
End Sub

Friend Sub SinkDOMContentLoaded()
    pvQueueEvent ev_DOMContentLoaded
End Sub

Friend Sub SinkStatusBarTextChanged()
    pvQueueEvent ev_StatusBarTextChanged, StatusBarText
End Sub

Friend Sub SinkContainsFullScreenElementChanged()
    pvQueueEvent ev_ContainsFullScreenElementChanged
End Sub

Friend Sub SinkZoomFactorChanged()
    pvQueueEvent ev_ZoomFactorChanged
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
        pvQueueEvent ev_GotFocus, m_eLastFocusReason
        m_eLastFocusReason = FocusReason_PROGRAMMATIC
    Else
        pvQueueEvent ev_LostFocus, FocusReason_PROGRAMMATIC
    End If
End Sub

Friend Sub SinkWebMessageReceived(ByVal bIsString As Boolean, sMsg As String, sJson As String)
    Dim oContent        As Collection
    Dim vResult         As Variant

    On Error GoTo EH
    If bIsString Then
        pvQueueEvent ev_JSMessage, sMsg, vbNullString, oContent
    Else
        pvJsonParseValue sJson, 1, vResult
        If IsObject(vResult) Then
            Set oContent = vResult
        End If
        pvQueueEvent ev_JSMessage, vbNullString, sJson, oContent
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
        pvQueueEvent ev_JSAsyncResult, vResult, cToken, sError
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

Friend Sub SinkBasicAuthenticationRequested(oArgs As IVBCoreWebView2BasicAuthenticationRequestedEventArgs)
    Dim oResponse       As IVBCoreWebView2BasicAuthenticationResponse
    Dim sUserName       As String
    Dim sPassword       As String
    Dim bCancel         As Boolean

    On Error GoTo EH
    bCancel = oArgs.Cancel
    RaiseEvent BasicAuthenticationRequested(pvToString(oArgs.URI), pvToString(oArgs.Challenge), sUserName, sPassword, bCancel)
    oArgs.Cancel = -bCancel
    If LenB(sUserName) <> 0 Or LenB(sPassword) <> 0 Then
        Set oResponse = oArgs.Response
        oResponse.UserName = StrPtr(sUserName)
        oResponse.Password = StrPtr(sPassword)
    End If
    Exit Sub
EH:
    PrintError "SinkBasicAuthenticationRequested"
End Sub

Friend Sub OnGetCookies(ByVal hResult As Long, oList As IVBCoreWebView2CookieList, ByVal cToken As Currency)
    Const SECS_PER_DAY                  As Long = 86400
    Dim cResult         As Collection
    Dim cItem           As Collection
    Dim oCookie         As IVBCoreWebView2Cookie
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
            Set oCookie = oList.GetValueAtIndex(lIdx)
            Set cItem = New Collection
            cItem.Add pvToString(oCookie.Name), "Name"
            cItem.Add pvToString(oCookie.Value), "Value"
            cItem.Add pvToString(oCookie.Domain), "Domain"
            cItem.Add pvToString(oCookie.Path), "Path"
            cItem.Add CBool(oCookie.IsSession), "IsSession"
            If Not oCookie.IsSession Then
                cItem.Add #1/1/1970# + oCookie.Expires / SECS_PER_DAY, "Expires"
            End If
            cItem.Add CBool(oCookie.IsSecure), "IsSecure"
            cItem.Add CBool(oCookie.IsHttpOnly), "IsHttpOnly"
            cItem.Add oCookie.SameSite, "SameSite"
            '--- no key on purpose: cookie names repeat across domains/paths
            cResult.Add cItem
        Next
    End If
    m_cScriptResults.Add Array(hResult, cResult), "T" & cToken
    Exit Sub
EH:
    PrintError "OnGetCookies"
End Sub

Friend Sub OnPrintToPdfStream(ByVal hResult As Long, oPdfStream As IVBStream, ByVal cToken As Currency)
    Const STREAM_SEEK_SET               As Long = 0
    Const STREAM_SEEK_END               As Long = 2
    Const E_FAIL                        As Long = &H80004005
    Dim baResult()      As Byte
    Dim cPos            As Currency
    Dim lSize           As Long
    Dim lRead           As Long

    On Error GoTo EH
    If Not pvSearchCollection(m_cScriptResults, "P" & cToken) Then
        '--- unknown token (timed-out or fire-and-forget call) - discard result
        Exit Sub
    End If
    m_cScriptResults.Remove "P" & cToken
    If hResult >= 0 And Not oPdfStream Is Nothing Then
        '--- IVBStream positions travel as CURRENCY i.e. scaled by 1/10000
        oPdfStream.Seek 0, STREAM_SEEK_END, cPos
        lSize = cPos * 10000
    End If
    If lSize > 0 Then
        oPdfStream.Seek 0, STREAM_SEEK_SET, cPos
        ReDim baResult(0 To lSize - 1)
        oPdfStream.Read VarPtr(baResult(0)), lSize, lRead
        m_cScriptResults.Add Array(hResult, baResult), "T" & cToken
    Else
        If hResult >= 0 Then
            hResult = E_FAIL
        End If
        m_cScriptResults.Add Array(hResult, vbNullString), "T" & cToken
    End If
    Exit Sub
EH:
    PrintError "OnPrintToPdfStream"
End Sub

'=========================================================================
' MST
'=========================================================================

Public Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgeogEQUAV1aLdCQUg8YIgz4AdCqL+oHHDBMFAIvCBSgSBQCri8IFZBIFAKuLwgV0EgUAqzPAq7kIAAAA86WBwgwTBQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItDBMFAAWgEgUAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g8EwUALSARBQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBItCEIXAdFuLCotBKIXAdCdS/9Bag/gBd0mLClL/USxahcB1PosKUmrw/3Eg/1EkWqkAAAAIdSuLClL/cghqAP9RHFr/QgQzwFBU/3IQ/1IUi1QkCMdCCAAAAABS6G////9YwhQADx8AjURAAQ==" ' 13.5.2020 18:59:12
    Const THUNK_SIZE    As Long = 5660
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function TerminateFireOnceTimerThunk(pTimer As IUnknown, pObj As Object) As IUnknown
    If Not pTimer Is Nothing Then
        Debug.Assert ThunkPrivateData(pTimer, 2) = ObjPtr(pObj)
        ThunkPrivateData(pTimer, 2) = 0
        Set pTimer = Nothing
    End If
End Function

Public Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, PTR_SIZE)
    End If
End Property

Public Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, lValue, PTR_SIZE)
    End If
End Property

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Const MEM_COMMIT                    As Long = &H1000
    Const PAGE_EXECUTE_READWRITE        As Long = &H40
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_dbljsCallTimeOutSeconds = 8
    Set m_cScriptResults = New Collection
    m_lDeferMsg = RegisterWindowMessage(StrPtr("cWebView2 Defer Event"))
End Sub

Private Sub Class_Terminate()
    Shutdown
End Sub
