VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BartNet Chat 3.0 - [Server]"
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
   Begin MSWinsockLib.Winsock wChat 
      Index           =   0
      Left            =   2880
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wListen 
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   2400
      Width           =   6495
   End
   Begin VB.TextBox txtChat 
      Height          =   2175
      Left            =   143
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "123.123.123.123"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   3030
      Width           =   2775
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.bartnet.freeservers.com [13132 user(s) online]"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   3030
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

Private Sub cmdSend_Click()
    Dim a As Integer
    Dim Message As String
    
    Message = "Server says: " & txtSend.Text
    
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then
        
        Else
            wChat(a).SendData Message
            DoEvents
        End If
        
        a = a + 1
    Loop
    
    txtChat.Text = txtChat.Text & Message & vbCrLf
    
    txtSend.Text = ""
End Sub

Private Sub Form_Load()
    wListen.LocalPort = 1
    wListen.Listen
    
    wChat(0).LocalPort = 2
    
    OnlineCount = 0
    
    cmdSend.Default = True
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    lblCredits.Caption = "www.bartnet.freeservers.com [" & OnlineCount & " user(s) online]"
    
    lblIP.Caption = "Your IP = " & wListen.LocalIP
End Sub

Private Sub wChat_Close(Index As Integer)
    wChat(Index).Close
    UpdateUserCount
End Sub

Private Sub wChat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As Integer
    Dim Message As String
    Dim Data As String
    
    wChat(Index).GetData Data, vbString
    
    Message = "Client " & Index & " says: " & Data
    
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then
        
        Else
            wChat(a).SendData Message
            DoEvents
        End If
        
        a = a + 1
    Loop
    
    txtChat.Text = txtChat.Text & Message & vbCrLf
End Sub


Private Sub wChat_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "An error has ocurred and te connection will therefore be broken." & vbCrLf & vbCrLf & "Error description : " & Description, vbOKOnly + vbCritical, "BartNet Chat 3.0"
    CloseAll
End Sub

Private Sub wListen_ConnectionRequest(ByVal requestID As Long)
    Dim a As Integer
    
    a = GetNextWinsock

    wChat(a).Close
    wChat(a).Accept requestID
    
    UpdateUserCount
End Sub

Private Function GetNextWinsock()
    Dim a As Integer
    
    a = wChat.Count

    Load wChat(a)
    wChat(a).LocalPort = a + 2
    
    GetNextWinsock = a
End Function

Private Sub UpdateUserCount()
    Dim a As Integer
    Dim Online As Integer
    
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then
        
        Else
            Online = Online + 1
        End If
        
        a = a + 1
    Loop
    
    OnlineCount = Online
    lblCredits.Caption = "www.bartnet.freeservers.com [" & OnlineCount & " user(s) online]"
    
    a = 0
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then

        Else
            wChat(a).SendData "Online" & OnlineCount
            DoEvents
        End If
        a = a + 1
    Loop
End Sub

Private Sub wListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "An error has ocurred and te connection will therefore be broken." & vbCrLf & vbCrLf & "Error description : " & Description, vbOKOnly + vbCritical, "BartNet Chat 3.0"
    CloseAll
End Sub

Private Sub CloseAll()
    wListen.Close
    
    Dim a As Integer
    
    Do Until a = wChat.Count
        wChat(a).Close
        a = a + 1
    Loop
    
    OnlineCount = 0
    
    lblCredits.Caption = "www.bartnet.freeservers.com [" & OnlineCount & " user(s) online]"
End Sub
