VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebView2Callback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' CLASS: cWebView2Callback
' Single class implementing every VB6-safe companion interface a cWebView2
' instance needs on the native side: the environment options bag, the
' one-shot completed-handlers (environment, controller, script, capture,
' devtools) and the long-lived event handlers. Each Invoke forwards to a
' Friend sink on the owning cWebView2, reached only through the pvOwner
' weak-reference accessor (m_lOwnerWeakRef holds a raw ObjPtr, never a real
' Set reference) so cWebView2 <-> long-lived callback instances never form
' a reference cycle; every Invoke bails out early once m_lOwnerWeakRef has
' been cleared by Terminate. One-shot instances carry a Currency token used
' for completion correlation; the same token discriminates the two
' GotFocus/LostFocus subscriptions which share a single native interface.
' Whichever role isn't in use for an instance is never invoked
'=========================================================================
Option Explicit
DefObj A-Z
Implements IVBWebView2EnvironmentOptions
Implements IVBWebView2EnvironmentOptions2
Implements IVBWebView2EnvironmentOptions3
Implements IVBWebView2EnvironmentOptions4
Implements IVBWebView2EnvironmentOptions5
Implements IVBWebView2EnvironmentOptions6
Implements IVBWebView2CreateEnvironmentCompletedHandler
Implements IVBWebView2CreateControllerCompletedHandler
Implements IVBWebView2NavigationStartingEventHandler
Implements IVBWebView2NavigationCompletedEventHandler
Implements IVBWebView2SourceChangedEventHandler
Implements IVBWebView2ProcessFailedEventHandler
Implements IVBWebView2ExecuteScriptCompletedHandler
Implements IVBWebView2CapturePreviewCompletedHandler
Implements IVBWebView2PrintToPdfCompletedHandler
Implements IVBWebView2ProfileAddBrowserExtensionCompletedHandler
Implements IVBWebView2ProfileGetBrowserExtensionsCompletedHandler
Implements IVBWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler
Implements IVBWebView2CallDevToolsProtocolMethodCompletedHandler
Implements IVBWebView2WebMessageReceivedEventHandler
Implements IVBWebView2ScriptDialogOpeningEventHandler
Implements IVBWebView2PermissionRequestedEventHandler
Implements IVBWebView2NewWindowRequestedEventHandler
Implements IVBWebView2AcceleratorKeyPressedEventHandler
Implements IVBWebView2WebResourceRequestedEventHandler
Implements IVBWebView2FocusChangedEventHandler
Implements IVBWebView2WebResourceResponseReceivedEventHandler
Implements IVBWebView2FrameCreatedEventHandler
Implements IVBWebView2DownloadStartingEventHandler
Implements IVBWebView2DocumentTitleChangedEventHandler
Implements IVBWebView2ContentLoadingEventHandler
Implements IVBWebView2HistoryChangedEventHandler
Implements IVBWebView2WindowCloseRequestedEventHandler
Implements IVBWebView2ZoomFactorChangedEventHandler
Implements IVBWebView2MoveFocusRequestedEventHandler
Implements IVBWebView2ContextMenuRequestedEventHandler

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, ByVal lpsz As Long) As Long
Private Declare Function GetAvailableCoreWebView2BrowserVersionString Lib "WebView2Loader" (ByVal browserExecutableFolder As Long, ByRef versionInfo As Long) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lOwnerWeakRef             As Long
Private m_cToken                    As Currency
Private m_sAdditionalBrowserArguments As String
Private m_sLanguage                 As String
Private m_sTargetCompatibleBrowserVersion As String
Private m_bAllowSSO                 As Boolean
Private m_bExclusiveUserDataFolderAccess As Boolean
Private m_bIsCustomCrashReportingEnabled As Boolean
Private m_bEnableTrackingPrevention As Boolean
Private m_bAreBrowserExtensionsEnabled As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Let AdditionalBrowserArguments(sValue As String)
    m_sAdditionalBrowserArguments = sValue
End Property

Public Property Let Language(sValue As String)
    m_sLanguage = sValue
End Property

Public Property Let TargetCompatibleBrowserVersion(sValue As String)
    m_sTargetCompatibleBrowserVersion = sValue
End Property

Public Property Let AllowSingleSignOnUsingOSPrimaryAccount(ByVal bValue As Boolean)
    m_bAllowSSO = bValue
End Property

Public Property Let ExclusiveUserDataFolderAccess(ByVal bValue As Boolean)
    m_bExclusiveUserDataFolderAccess = bValue
End Property

Public Property Let IsCustomCrashReportingEnabled(ByVal bValue As Boolean)
    m_bIsCustomCrashReportingEnabled = bValue
End Property

Public Property Let EnableTrackingPrevention(ByVal bValue As Boolean)
    m_bEnableTrackingPrevention = bValue
End Property

Public Property Let AreBrowserExtensionsEnabled(ByVal bValue As Boolean)
    m_bAreBrowserExtensionsEnabled = bValue
End Property

Private Property Get pvOwner() As cWebView2
    Call vbaObjSetAddref(pvOwner, m_lOwnerWeakRef)
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Sub Init(oOwner As cWebView2, Optional ByVal cToken As Currency)
    m_lOwnerWeakRef = ObjPtr(oOwner)
    m_cToken = cToken
End Sub

Friend Sub Terminate()
    m_lOwnerWeakRef = 0
End Sub

Public Sub RaiseContextMenuEvent(ByVal ScreenX As Variant, ByVal ScreenY As Variant)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.RaiseContextMenuEvent ScreenX, ScreenY
End Sub

Public Sub RaiseMessageEvent(sMsg As Variant, sMsgContent As Variant)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.RaiseMessageEvent sMsg, sMsgContent
End Sub

'= private ===============================================================

Private Function pvToString(ByVal lPtr As Long, Optional ByVal DontFree As Boolean) As String
    Call SysReAllocString(VarPtr(pvToString), lPtr)
    If Not DontFree Then
        Call CoTaskMemFree(lPtr)
    End If
End Function

Private Function pvFromString(sValue As String) As Long
    If StrPtr(sValue) <> 0 Then
        pvFromString = CoTaskMemAlloc(LenB(sValue) + 2)
        Call CopyMemory(ByVal pvFromString, ByVal StrPtr(sValue), LenB(sValue) + 2)
    End If
End Function

'=========================================================================
' IVBWebView2EnvironmentOptions interface
'=========================================================================

Private Property Get IVBWebView2EnvironmentOptions_AdditionalBrowserArguments() As Long
    IVBWebView2EnvironmentOptions_AdditionalBrowserArguments = pvFromString(m_sAdditionalBrowserArguments)
End Property

Private Property Let IVBWebView2EnvironmentOptions_AdditionalBrowserArguments(ByVal RHS As Long)
    m_sAdditionalBrowserArguments = pvToString(RHS, DontFree:=True)
End Property

Private Property Get IVBWebView2EnvironmentOptions_Language() As Long
    IVBWebView2EnvironmentOptions_Language = pvFromString(m_sLanguage)
End Property

Private Property Let IVBWebView2EnvironmentOptions_Language(ByVal RHS As Long)
    m_sLanguage = pvToString(RHS, DontFree:=True)
End Property

Private Property Get IVBWebView2EnvironmentOptions_TargetCompatibleBrowserVersion() As Long
    Dim lVersionPtr     As Long

    If LenB(m_sTargetCompatibleBrowserVersion) = 0 Then
        Call GetAvailableCoreWebView2BrowserVersionString(0, lVersionPtr)
        IVBWebView2EnvironmentOptions_TargetCompatibleBrowserVersion = lVersionPtr
    Else
        IVBWebView2EnvironmentOptions_TargetCompatibleBrowserVersion = pvFromString(m_sTargetCompatibleBrowserVersion)
    End If
End Property

Private Property Let IVBWebView2EnvironmentOptions_TargetCompatibleBrowserVersion(ByVal RHS As Long)
    m_sTargetCompatibleBrowserVersion = pvToString(RHS, DontFree:=True)
End Property

Private Property Get IVBWebView2EnvironmentOptions_AllowSingleSignOnUsingOSPrimaryAccount() As Long
    IVBWebView2EnvironmentOptions_AllowSingleSignOnUsingOSPrimaryAccount = -m_bAllowSSO
End Property

Private Property Let IVBWebView2EnvironmentOptions_AllowSingleSignOnUsingOSPrimaryAccount(ByVal RHS As Long)
    m_bAllowSSO = (RHS <> 0)
End Property

Private Property Get IVBWebView2EnvironmentOptions2_ExclusiveUserDataFolderAccess() As Long
    IVBWebView2EnvironmentOptions2_ExclusiveUserDataFolderAccess = -m_bExclusiveUserDataFolderAccess
End Property

Private Property Let IVBWebView2EnvironmentOptions2_ExclusiveUserDataFolderAccess(ByVal RHS As Long)
    m_bExclusiveUserDataFolderAccess = (RHS <> 0)
End Property

Private Property Get IVBWebView2EnvironmentOptions3_IsCustomCrashReportingEnabled() As Long
    IVBWebView2EnvironmentOptions3_IsCustomCrashReportingEnabled = -m_bIsCustomCrashReportingEnabled
End Property

Private Property Let IVBWebView2EnvironmentOptions3_IsCustomCrashReportingEnabled(ByVal RHS As Long)
    m_bIsCustomCrashReportingEnabled = (RHS <> 0)
End Property

Private Sub IVBWebView2EnvironmentOptions4_GetCustomSchemeRegistrations(lCount As Long, lSchemeRegistrations As Long)
    '--- custom scheme registration is out of scope (no RC6 counterpart) -
    '--- always report zero registrations
    lCount = 0
    lSchemeRegistrations = 0
End Sub

Private Sub IVBWebView2EnvironmentOptions4_SetCustomSchemeRegistrations(ByVal lCount As Long, ByVal lSchemeRegistrations As Long)
End Sub

Private Property Get IVBWebView2EnvironmentOptions5_EnableTrackingPrevention() As Long
    IVBWebView2EnvironmentOptions5_EnableTrackingPrevention = -m_bEnableTrackingPrevention
End Property

Private Property Let IVBWebView2EnvironmentOptions5_EnableTrackingPrevention(ByVal RHS As Long)
    m_bEnableTrackingPrevention = (RHS <> 0)
End Property

Private Property Get IVBWebView2EnvironmentOptions6_AreBrowserExtensionsEnabled() As Long
    IVBWebView2EnvironmentOptions6_AreBrowserExtensionsEnabled = -m_bAreBrowserExtensionsEnabled
End Property

Private Property Let IVBWebView2EnvironmentOptions6_AreBrowserExtensionsEnabled(ByVal RHS As Long)
    m_bAreBrowserExtensionsEnabled = (RHS <> 0)
End Property

'=========================================================================
' IVBWebView2CreateEnvironmentCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2CreateEnvironmentCompletedHandler_Invoke(ByVal hResult As Long, ByVal createdEnvironment As IVBCoreWebView2Environment)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.OnEnvironmentCreated hResult, createdEnvironment
End Sub

'=========================================================================
' IVBWebView2CreateControllerCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2CreateControllerCompletedHandler_Invoke(ByVal hResult As Long, ByVal createdController As IVBCoreWebView2Controller)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.OnControllerCreated hResult, createdController
End Sub

'=========================================================================
' IVBWebView2NavigationStartingEventHandler interface
'=========================================================================

Private Sub IVBWebView2NavigationStartingEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2NavigationStartingEventArgs)
    Dim bCancel         As Boolean

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    bCancel = oArgs.Cancel
    pvOwner.SinkNavigationStarting oArgs.IsUserInitiated, oArgs.IsRedirected, pvToString(oArgs.URI), bCancel
    oArgs.Cancel = -bCancel
End Sub

'=========================================================================
' IVBWebView2NavigationCompletedEventHandler interface
'=========================================================================

Private Sub IVBWebView2NavigationCompletedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2NavigationCompletedEventArgs)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkNavigationCompleted oArgs.IsSuccess, oArgs.WebErrorStatus
End Sub

'=========================================================================
' IVBWebView2SourceChangedEventHandler interface
'=========================================================================

Private Sub IVBWebView2SourceChangedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal args As VB6IUnknown)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkSourceChanged
End Sub

'=========================================================================
' IVBWebView2ProcessFailedEventHandler interface
'=========================================================================

Private Sub IVBWebView2ProcessFailedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal args As VB6IUnknown)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkProcessFailed
End Sub

'=========================================================================
' IVBWebView2ExecuteScriptCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2ExecuteScriptCompletedHandler_Invoke(ByVal hResult As Long, ByVal resultObjectAsJson As Long)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.OnExecuteScriptCompleted hResult, pvToString(resultObjectAsJson, DontFree:=True), m_cToken
End Sub

'=========================================================================
' IVBWebView2CapturePreviewCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2CapturePreviewCompletedHandler_Invoke(ByVal hResult As Long)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.OnExecuteScriptCompleted hResult, vbNullString, m_cToken
End Sub

'=========================================================================
' IVBWebView2PrintToPdfCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2PrintToPdfCompletedHandler_Invoke(ByVal hResult As Long, ByVal lIsSuccessful As Long)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    '--- payload is drained untouched by the sync waiter so "true"/"false"
    '--- ride the same channel as script results
    If lIsSuccessful <> 0 Then
        pvOwner.OnExecuteScriptCompleted hResult, "true", m_cToken
    Else
        pvOwner.OnExecuteScriptCompleted hResult, "false", m_cToken
    End If
End Sub

'=========================================================================
' IVBWebView2ProfileAddBrowserExtensionCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2ProfileAddBrowserExtensionCompletedHandler_Invoke(ByVal hResult As Long, ByVal oResult As IVBCoreWebView2BrowserExtension)
    Dim sId             As String

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    If Not oResult Is Nothing Then
        sId = pvToString(oResult.Id)
    End If
    pvOwner.OnExecuteScriptCompleted hResult, sId, m_cToken
End Sub

'=========================================================================
' IVBWebView2ProfileGetBrowserExtensionsCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2ProfileGetBrowserExtensionsCompletedHandler_Invoke(ByVal hResult As Long, ByVal oResult As IVBCoreWebView2BrowserExtensionList)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.OnGetBrowserExtensions hResult, oResult, m_cToken
End Sub

'=========================================================================
' IVBWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke(ByVal hResult As Long, ByVal id As Long)
    '--- fire-and-forget - script id not retained (RemoveScriptToExecuteOnDocumentCreated not in RC6 API)
End Sub

'=========================================================================
' IVBWebView2CallDevToolsProtocolMethodCompletedHandler interface
'=========================================================================

Private Sub IVBWebView2CallDevToolsProtocolMethodCompletedHandler_Invoke(ByVal hResult As Long, ByVal returnObjectAsJson As Long)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.OnExecuteScriptCompleted hResult, pvToString(returnObjectAsJson, DontFree:=True), m_cToken
End Sub

'=========================================================================
' IVBWebView2WebMessageReceivedEventHandler interface
'=========================================================================

Private Sub IVBWebView2WebMessageReceivedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2WebMessageReceivedEventArgs)
    Dim lPtr            As Long
    Dim bIsString       As Boolean
    Dim sMsg            As String

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    If oArgs.TryGetWebMessageAsString(lPtr) >= 0 Then
        sMsg = pvToString(lPtr)
        bIsString = True
    End If
    pvOwner.SinkWebMessageReceived bIsString, sMsg, pvToString(oArgs.WebMessageAsJson)
End Sub

'=========================================================================
' IVBWebView2ScriptDialogOpeningEventHandler interface
'=========================================================================

Private Sub IVBWebView2ScriptDialogOpeningEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2ScriptDialogOpeningEventArgs)
    Dim bAccept         As Boolean
    Dim sResultText     As String

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkScriptDialogOpening oArgs.Kind, bAccept, sResultText, pvToString(oArgs.URI), pvToString(oArgs.Message), pvToString(oArgs.DefaultText)
    If bAccept Then
        oArgs.ResultText = StrPtr(sResultText)
        oArgs.Accept
    End If
End Sub

'=========================================================================
' IVBWebView2PermissionRequestedEventHandler interface
'=========================================================================

Private Sub IVBWebView2PermissionRequestedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2PermissionRequestedEventArgs)
    Dim eState          As Long

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    eState = oArgs.State
    pvOwner.SinkPermissionRequested oArgs.IsUserInitiated, eState, pvToString(oArgs.URI), oArgs.PermissionKind
    oArgs.State = eState
End Sub

'=========================================================================
' IVBWebView2NewWindowRequestedEventHandler interface
'=========================================================================

Private Sub IVBWebView2NewWindowRequestedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2NewWindowRequestedEventArgs)
    Dim oFeatures       As IVBCoreWebView2WindowFeatures
    Dim cFeatures       As Collection
    Dim bHandled        As Boolean

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    Set oFeatures = oArgs.WindowFeatures
    Set cFeatures = New Collection
    cFeatures.Add CBool(oFeatures.HasPosition), "HasPosition"
    cFeatures.Add CBool(oFeatures.HasSize), "HasSize"
    cFeatures.Add oFeatures.Left, "Left"
    cFeatures.Add oFeatures.Top, "Top"
    cFeatures.Add oFeatures.Width, "Width"
    cFeatures.Add oFeatures.Height, "Height"
    cFeatures.Add CBool(oFeatures.ShouldDisplayMenuBar), "ShouldDisplayMenuBar"
    cFeatures.Add CBool(oFeatures.ShouldDisplayStatus), "ShouldDisplayStatus"
    cFeatures.Add CBool(oFeatures.ShouldDisplayToolbar), "ShouldDisplayToolbar"
    cFeatures.Add CBool(oFeatures.ShouldDisplayScrollBars), "ShouldDisplayScrollBars"
    bHandled = oArgs.Handled
    pvOwner.SinkNewWindowRequested oArgs.IsUserInitiated, bHandled, pvToString(oArgs.URI), cFeatures
    oArgs.Handled = -bHandled
End Sub

'=========================================================================
' IVBWebView2AcceleratorKeyPressedEventHandler interface
'=========================================================================

Private Sub IVBWebView2AcceleratorKeyPressedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2AcceleratorKeyPressedEventArgs)
    Dim lParam          As Long
    Dim bHandled        As Boolean

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    lParam = oArgs.KeyEventLParam
    bHandled = oArgs.Handled
    '--- decode WM_KEYDOWN lParam bits: 0-15 repeat, 16-23 scan code,
    '--- 24 extended, 29 context (menu), 30 previous state, 31 transition
    pvOwner.SinkAcceleratorKeyPressed oArgs.KeyEventKind, _
        (lParam And &H1000000) <> 0, _
        (lParam And &H40000000) <> 0, _
        lParam < 0, _
        (lParam And &H20000000) <> 0, _
        lParam And &HFFFF&, _
        (lParam \ &H10000) And &HFF&, _
        bHandled
    oArgs.Handled = -bHandled
End Sub

'=========================================================================
' IVBWebView2WebResourceRequestedEventHandler interface
'=========================================================================

Private Sub IVBWebView2WebResourceRequestedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2WebResourceRequestedEventArgs)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkWebResourceRequested oArgs
End Sub

'=========================================================================
' IVBWebView2FocusChangedEventHandler interface
'=========================================================================

Private Sub IVBWebView2FocusChangedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal args As VB6IUnknown)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkFocusChanged m_cToken = 1
End Sub

'=========================================================================
' IVBWebView2WebResourceResponseReceivedEventHandler interface
'=========================================================================

Private Sub IVBWebView2WebResourceResponseReceivedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2WebResourceResponseReceivedEventArgs)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkWebResourceResponseReceived oArgs
End Sub

'=========================================================================
' IVBWebView2FrameCreatedEventHandler interface
'=========================================================================

Private Sub IVBWebView2FrameCreatedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2FrameCreatedEventArgs)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkFrameCreated pvToString(oArgs.Frame.Name)
End Sub

'=========================================================================
' IVBWebView2DownloadStartingEventHandler interface
'=========================================================================

Private Sub IVBWebView2DownloadStartingEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2DownloadStartingEventArgs)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkDownloadStarting oArgs
End Sub

'=========================================================================
' IVBWebView2DocumentTitleChangedEventHandler interface
'=========================================================================

Private Sub IVBWebView2DocumentTitleChangedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal args As VB6IUnknown)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkDocumentTitleChanged
End Sub

'=========================================================================
' IVBWebView2ContentLoadingEventHandler interface
'=========================================================================

Private Sub IVBWebView2ContentLoadingEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2ContentLoadingEventArgs)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkContentLoading (oArgs.IsErrorPage <> 0)
End Sub

'=========================================================================
' IVBWebView2HistoryChangedEventHandler interface
'=========================================================================

Private Sub IVBWebView2HistoryChangedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal args As VB6IUnknown)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkHistoryChanged
End Sub

'=========================================================================
' IVBWebView2WindowCloseRequestedEventHandler interface
'=========================================================================

Private Sub IVBWebView2WindowCloseRequestedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal args As VB6IUnknown)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkWindowCloseRequested
End Sub

'=========================================================================
' IVBWebView2ZoomFactorChangedEventHandler interface
'=========================================================================

Private Sub IVBWebView2ZoomFactorChangedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal args As VB6IUnknown)
    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    pvOwner.SinkZoomFactorChanged
End Sub

'=========================================================================
' IVBWebView2MoveFocusRequestedEventHandler interface
'=========================================================================

Private Sub IVBWebView2MoveFocusRequestedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2MoveFocusRequestedEventArgs)
    Dim bHandled        As Boolean

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    bHandled = oArgs.Handled
    pvOwner.SinkMoveFocusRequested oArgs.Reason, bHandled
    oArgs.Handled = -bHandled
End Sub

'=========================================================================
' IVBWebView2ContextMenuRequestedEventHandler interface
'=========================================================================

Private Sub IVBWebView2ContextMenuRequestedEventHandler_Invoke(ByVal sender As VB6IUnknown, ByVal oArgs As IVBCoreWebView2ContextMenuRequestedEventArgs)
    Dim oTarget         As IVBCoreWebView2ContextMenuTarget
    Dim uLocation       As APIPOINT
    Dim sLinkURI        As String
    Dim sSelectionText  As String
    Dim bHandled        As Boolean

    If m_lOwnerWeakRef = 0 Then
        Exit Sub
    End If
    Set oTarget = oArgs.ContextMenuTarget
    If oTarget.HasLinkUri <> 0 Then
        sLinkURI = pvToString(oTarget.LinkURI)
    End If
    If oTarget.HasSelection <> 0 Then
        sSelectionText = pvToString(oTarget.SelectionText)
    End If
    uLocation = oArgs.Location
    bHandled = oArgs.Handled
    pvOwner.SinkContextMenuRequested pvToString(oTarget.PageURI), sLinkURI, sSelectionText, uLocation.x, uLocation.y, bHandled
    oArgs.Handled = -bHandled
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_bEnableTrackingPrevention = True
End Sub
