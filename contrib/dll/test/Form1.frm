VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8892
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8892
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picVW 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   6648
      Left            =   168
      ScaleHeight     =   6648
      ScaleWidth      =   10344
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   10344
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oVW As cWebView2
Attribute m_oVW.VB_VarHelpID = -1

Private Sub Form_Load()
    Visible = True
    Set m_oVW = New cWebView2
    m_oVW.BindTo picVW.hWnd
    m_oVW.Navigate "https://vbforums.com", 0
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        picVW.Move 0, picVW.Top, ScaleWidth, ScaleHeight - picVW.Top
        If Not m_oVW Is Nothing Then
            m_oVW.SyncSizeToHostWindow
        End If
    End If
End Sub

