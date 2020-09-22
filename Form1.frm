VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BartNet Chat 3.0 - [Client]"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   3390
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock w1 
      Left            =   1920
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Project1.chameleonButton cmdSend 
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   2410
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "S&end"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":3723
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtSend 
      Height          =   525
      Left            =   143
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   6495
   End
   Begin VB.TextBox txtChat 
      Height          =   2175
      Left            =   143
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   7215
   End
   Begin Project1.chameleonButton cmdConnect 
      Height          =   285
      Left            =   6503
      TabIndex        =   1
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "C&onnect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":373F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtConnect 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4583
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.bartnet.freeservers.com [13132 user(s) online]"
      Height          =   255
      Left            =   143
      TabIndex        =   4
      Top             =   3035
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
'' Created By BelgiumBoy_007    ''
''                              ''
'' Copyright 2002 BartNet corp. ''
''                              ''
'' www.bartnet.freeservers.com  ''
''''''''''''''''''''''''''''''''''

Private OnlineCount As Integer

Private Sub cmdConnect_Click()
    If Len(txtConnect.Text) = 0 Then
    
    Else
        w1.Close
        w1.Connect txtConnect.Text, 1
    End If
End Sub

Private Sub cmdSend_Click()
    If Len(txtSend.Text) = 0 Then
    
    Else
        w1.SendData txtSend.Text
        txtSend.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Status 2
    
    cmdSend.Default = True
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Status(ByVal WhichStatus As Integer)
    Select Case WhichStatus
        Case 1 'Connected
            txtChat.Enabled = True
            txtSend.Enabled = True
            cmdSend.Enabled = True
            lblCredits.Caption = "www.bartnet.freeservers.com [" & OnlineCount & " user(s) online]"
            cmdConnect.Enabled = False
            txtConnect.Enabled = False
        Case 2 'Disconnected
            txtChat.Enabled = False
            txtSend.Enabled = False
            cmdSend.Enabled = False
            lblCredits.Caption = "www.bartnet.freeservers.com"
            cmdConnect.Enabled = True
            txtConnect.Enabled = True
            w1.Close
    End Select
End Sub

Private Sub w1_Close()
    Status 2
End Sub

Private Sub w1_Connect()
    Status 1
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    
    w1.GetData data, vbString
    
    If Mid(data, 1, 6) = "Online" Then
        OnlineCount = Mid(data, 7, Len(data) - 6)
        lblCredits.Caption = "www.bartnet.freeservers.com [" & OnlineCount & " user(s) online]"
    Else
        txtChat.Text = txtChat.Text & data & vbCrLf
    End If
End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "An error has ocurred and te connection will therefore be broken." & vbCrLf & vbCrLf & "Error description : " & Description, vbOKOnly + vbCritical, "BartNet Chat 3.0"
    w1.Close
End Sub
