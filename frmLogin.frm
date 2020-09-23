VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WhiteBoardPro Login"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2100
      Top             =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Server"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Anonymous"
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Waiting..."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   4950
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   5040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   2640
      X2              =   2640
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your name below, and then chose to be the server or to connect."
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.ws.LocalPort = 6050
frmMain.ws.Listen
lblStatus.Caption = "Status: Awaiting Connection"
Text2.Enabled = False
Command2.Enabled = False
End Sub

Private Sub Command2_Click()
frmMain.ws.RemoteHost = Text2.Text
frmMain.ws.RemotePort = 6050
frmMain.ws.Connect
lblStatus.Caption = "Status: Connecting..."
Command1.Enabled = True

TimeOut.Enabled = True

End Sub

Private Sub Form_Load()
Load frmMain
Text2.Text = frmMain.ws.LocalIP
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMain
End Sub

Private Sub TimeOut_Timer()
TimeOut.Enabled = False
Text2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
lblStatus.Caption = "Status: Connection Attempt Timed Out."
frmMain.ws.Close
End Sub
