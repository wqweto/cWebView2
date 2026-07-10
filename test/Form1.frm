VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "cWebView2 Test Harness"
   ClientHeight    =   11568
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15528
   LinkTopic       =   "Form1"
   ScaleHeight     =   11568
   ScaleWidth      =   15528
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHost 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4692
      ScaleWidth      =   7920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   7920
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">"
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   315
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Width           =   840
   End
   Begin VB.TextBox txtUrl 
      Height          =   315
      Left            =   1980
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.TextBox txtJs 
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Text            =   "document.title"
      Top             =   360
      Width           =   6555
   End
   Begin VB.CommandButton cmdRunJs
      Caption         =   "Run JS"
      Height          =   315
      Left            =   6600
      TabIndex        =   7
      Top             =   360
      Width           =   840
   End
   Begin VB.CommandButton cmdPdf
      Caption         =   "PDF"
      Height          =   315
      Left            =   7500
      TabIndex        =   8
      Top             =   360
      Width           =   840
   End
   Begin VB.CommandButton cmdVHost
      Caption         =   "VHost"
      Height          =   315
      Left            =   6540
      TabIndex        =   9
      Top             =   360
      Width           =   840
   End
   Begin VB.CommandButton cmdDark
      Caption         =   "Dark"
      Height          =   315
      Left            =   5580
      TabIndex        =   10
      Top             =   360
      Width           =   840
   End
   Begin VB.Label lblStatus 
      Height          =   2400
      Left            =   0
      TabIndex        =   4
      Top             =   5085
      Width           =   7920
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' FORM: Form1
' Manual test harness for cWebView2, grown phase by phase
'=========================================================================

Option Explicit

Private WithEvents m_oWebView2          As cWebView2
Attribute m_oWebView2.VB_VarHelpID = -1

Private Sub Form_Load()
    Dim hResult         As Long

    Visible = True
    Set m_oWebView2 = New cWebView2
    hResult = m_oWebView2.BindTo(picHost.hWnd, , , , "--allow-run-as-system", , False, True, False)
    If hResult <> 0 Then
        MsgBox "BindTo failed: 0x" & Hex$(hResult), vbCritical
        Exit Sub
    End If
    m_oWebView2.AddScriptToExecuteOnDocumentCreated "var observer = new MutationObserver(function() { vbH().RaiseMessageEvent('title_change', document.title); });" & vbCrLf & _
        "document.addEventListener('DOMContentLoaded', function() { observer.observe(document.querySelector('title'), { childList: true }); vbH().RaiseMessageEvent('title_change', document.title); }); "
    m_oWebView2.AddScriptToExecuteOnDocumentCreated "chrome.webview.addEventListener('message', function(e) { vbH().RaiseMessageEvent('host_msg', '' + e.data); });"
    txtUrl.Text = "https://www.dir.bg"
    Call m_oWebView2.Navigate(txtUrl.Text)
    m_oWebView2.PostWebMessageAsString "hello-from-host"
    lblStatus.Caption = lblStatus.Caption & " | ver=" & m_oWebView2.BrowserVersion & " profile=" & m_oWebView2.ProfileName
    lblStatus.Caption = lblStatus.Caption & "jsRun(eval,40+2)=" & m_oWebView2.jsRun("eval", "40+2") & _
        ", jsRunAsync#" & m_oWebView2.jsRunAsync("eval", "6*7") & _
        ", UA.len=" & Len(m_oWebView2.UserAgent) & _
        ", pinch=" & m_oWebView2.IsPinchZoomEnabled
    m_oWebView2.AddObject "tst", Me
    m_oWebView2.ExecuteScript "chrome.webview.postMessage('ping=' + tst.Ping('42'))"
    m_oWebView2.ExecuteScript "chrome.webview.postMessage({answer:42})"
    m_oWebView2.AreDefaultScriptDialogsEnabled = False
    m_oWebView2.ExecuteScript "alert('dialog test')"
    Dim baPng()         As Byte
    baPng = m_oWebView2.CapturePreview
    On Error Resume Next
    lblStatus.Caption = lblStatus.Caption & " | png=" & (UBound(baPng) + 1) & " bytes"
    On Error GoTo 0
    Dim oCdp            As Object

    Set oCdp = m_oWebView2.CallDevToolsProtocolMethod("Browser.getVersion", "{}")
    If Not oCdp Is Nothing Then
        lblStatus.Caption = lblStatus.Caption & " | cdp.product=" & oCdp("product")
    End If
    m_oWebView2.ZoomFactor = 1.25
    lblStatus.Caption = lblStatus.Caption & " | Zoom=" & m_oWebView2.ZoomFactor
End Sub

Public Function Ping(ByVal sValue As String) As String
    Ping = "pong:" & sValue
End Function

Private Sub Form_Resize()
    If ScaleHeight <= 0 Or ScaleWidth <= 0 Then
        Exit Sub
    End If
    picHost.Move 0, txtJs.Top + txtJs.Height, ScaleWidth, ScaleHeight - txtJs.Top - txtJs.Height - lblStatus.Height
    lblStatus.Move 0, ScaleHeight - lblStatus.Height, ScaleWidth
    cmdGo.Move ScaleWidth - cmdGo.Width, 0
    txtUrl.Move txtUrl.Left, 0, cmdGo.Left - txtUrl.Left - 60
    cmdPdf.Move ScaleWidth - cmdPdf.Width, txtJs.Top
    cmdRunJs.Move cmdPdf.Left - cmdRunJs.Width - 60, txtJs.Top
    cmdVHost.Move cmdRunJs.Left - cmdVHost.Width - 60, txtJs.Top
    cmdDark.Move cmdVHost.Left - cmdDark.Width - 60, txtJs.Top
    txtJs.Move 0, txtJs.Top, cmdDark.Left - 60
    If Not m_oWebView2 Is Nothing Then
        Call m_oWebView2.SyncSizeToHostWindow
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_oWebView2 Is Nothing Then
        Call m_oWebView2.Shutdown
        Set m_oWebView2 = Nothing
    End If
End Sub

Private Sub cmdGo_Click()
    Call m_oWebView2.Navigate(txtUrl.Text)
End Sub

Private Sub cmdBack_Click()
    Call m_oWebView2.GoBack
End Sub

Private Sub cmdForward_Click()
    Call m_oWebView2.GoForward
End Sub

Private Sub txtUrl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call m_oWebView2.Navigate(txtUrl.Text)
    End If
End Sub

Private Sub m_oWebView2_InitComplete()
    lblStatus.Caption = "Ready"
End Sub

Private Sub m_oWebView2_NavigationStarting(ByVal IsUserInitiated As Boolean, ByVal IsRedirected As Boolean, ByVal URI As String, ByRef Cancel As Boolean)
    lblStatus.Caption = "Navigating to " & URI & " ..."
End Sub

Private Sub m_oWebView2_ContentLoading(ByVal IsErrorPage As Boolean)
    lblStatus.Caption = lblStatus.Caption & " | ContentLoading(err=" & IsErrorPage & ")"
End Sub

Private Sub m_oWebView2_DOMContentLoaded()
    lblStatus.Caption = lblStatus.Caption & " | DOMContentLoaded"
End Sub

Private Sub m_oWebView2_HistoryChanged()
    lblStatus.Caption = lblStatus.Caption & " | HistoryChanged(back=" & m_oWebView2.CanGoBack & ")"
End Sub

Private Sub m_oWebView2_ZoomFactorChanged()
    lblStatus.Caption = lblStatus.Caption & " | ZoomFactorChanged=" & m_oWebView2.ZoomFactor
End Sub

Private Sub m_oWebView2_ContextMenuRequested(ByVal PageURI As String, ByVal LinkURI As String, ByVal SelectionText As String, ByVal ScreenX As Long, ByVal ScreenY As Long, Handled As Boolean)
    lblStatus.Caption = lblStatus.Caption & " | ContextMenuRequested(" & ScreenX & "," & ScreenY & ")"
End Sub

Private Sub m_oWebView2_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
    txtUrl.Text = m_oWebView2.DocumentURL
    lblStatus.Caption = lblStatus.Caption & m_oWebView2.DocumentTitle & "  -  " & m_oWebView2.DocumentURL
    If Not IsSuccess Then
        lblStatus.Caption = lblStatus.Caption & "  (WebErrorStatus=" & WebErrorStatus & ")"
    End If
End Sub

Private Sub m_oWebView2_ProcessFailed()
    lblStatus.Caption = "Browser process failed"
End Sub

Private Sub m_oWebView2_JSAsyncResult(Result As Variant, ByVal Token As Currency, ByVal ErrString As String)
    lblStatus.Caption = lblStatus.Caption & " -> async#" & Token & "=" & Result & " " & ErrString
End Sub

Private Sub m_oWebView2_TitleChange(ByVal Text As String)
    lblStatus.Caption = lblStatus.Caption & " | TitleChange=" & Text
End Sub

Private Sub m_oWebView2_JSMessage(ByVal sMsg As String, ByVal sMsgContent As String, oJSONContent As Collection)
    If sMsg = "title_change" Then
        Caption = "cWebView2 Test Harness - " & sMsgContent
    ElseIf Not oJSONContent Is Nothing Then
        lblStatus.Caption = lblStatus.Caption & " | json.answer=" & oJSONContent("answer")
    ElseIf LenB(sMsgContent) <> 0 Then
        lblStatus.Caption = lblStatus.Caption & " | " & sMsg & "=" & sMsgContent
    Else
        lblStatus.Caption = lblStatus.Caption & " | msg=" & sMsg
    End If
End Sub

Private Sub cmdRunJs_Click()
    Dim vResult         As Variant

    On Error Resume Next
    vResult = m_oWebView2.jsProp(txtJs.Text)
    If LenB(m_oWebView2.jsLastError) <> 0 Then
        lblStatus.Caption = "Error: " & m_oWebView2.jsLastError
    Else
        lblStatus.Caption = "Result: " & vResult
    End If
End Sub

Private Sub txtJs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdRunJs_Click
    End If
End Sub

Private Sub cmdPdf_Click()
    Dim sFile           As String
    Dim vResult         As Variant

    sFile = App.Path & "\PrintToPdf.pdf"
    vResult = m_oWebView2.PrintToPdf(sFile)
    If vResult = True Then
        lblStatus.Caption = "Saved " & sFile & " (" & FileLen(sFile) & " bytes)"
    Else
        lblStatus.Caption = "PrintToPdf failed"
    End If
End Sub

Private Sub cmdDark_Click()
    If m_oWebView2.PreferredColorScheme = ColorScheme_DARK Then
        m_oWebView2.PreferredColorScheme = ColorScheme_LIGHT
    Else
        m_oWebView2.PreferredColorScheme = ColorScheme_DARK
    End If
    lblStatus.Caption = "PreferredColorScheme=" & m_oWebView2.PreferredColorScheme & " | IsMuted=" & m_oWebView2.IsMuted & " | StatusBarText=" & m_oWebView2.StatusBarText
End Sub

Private Sub cmdVHost_Click()
    Dim sDir            As String
    Dim sStatus         As String
    Dim vResult         As Variant

    On Error Resume Next
    sDir = App.Path & "\vhost"
    MkDir sDir
    Open sDir & "\index.html" For Output As #1
    Print #1, "<html><head><title>VHostOK</title></head><body><h1>Virtual host mapping works</h1></body></html>"
    Close #1
    m_oWebView2.SetVirtualHostNameToFolderMapping "appassets", sDir
    vResult = m_oWebView2.Navigate("https://appassets/index.html")
    sStatus = "vhost.Navigate=" & vResult & " title=" & m_oWebView2.DocumentTitle
    m_oWebView2.ClearVirtualHostNameToFolderMapping "appassets"
    vResult = m_oWebView2.Navigate("https://appassets/index.html")
    lblStatus.Caption = sStatus & " | cleared.Navigate=" & vResult & " (expected False)"
End Sub

Private Sub m_oWebView2_ScriptDialogOpening(ByVal ScriptDialogKind As eWebView2ScriptDialogKind, ByRef Accept As Boolean, ByRef ResultText As String, ByVal URI As String, ByVal Message As String, ByVal DefaultText As String)
    lblStatus.Caption = lblStatus.Caption & " | dialog(" & ScriptDialogKind & ")=" & Message
    Accept = True
End Sub

Private Sub m_oWebView2_UserContextMenu(ByVal ScreenX As Long, ByVal ScreenY As Long)
    lblStatus.Caption = lblStatus.Caption & " | ctxmenu@" & ScreenX & "," & ScreenY
End Sub

