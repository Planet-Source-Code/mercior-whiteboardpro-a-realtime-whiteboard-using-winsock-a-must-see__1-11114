VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WhiteBoardPro By Paul Blower"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save Picture"
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Speak"
      Default         =   -1  'True
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Text            =   "Type Chat Here"
      Top             =   7080
      Width           =   5415
   End
   Begin RichTextLib.RichTextBox CBox 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TransTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   1560
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Color"
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ______________________________________________________
' /                                                      \
' | Another (rather cool) bit of internet code here.     |
' | The top box is a whiteboard which both sides can     |
' | Draw onto. The display is updated every 1/2 second.  |
' | Last time, i got loads of comments on how cool my    |
' | code was, but only 2 votes! :(, so come on.. PLEASE  |
' | just take the time to vote for this..! (ill upload   |
' | the URL of my code here:                             |
' | http://www.netprogrammer.org/code.html, so you dont  |
' | have to go hunting for it again... arent i nice?)    |
' \______________________________________________________/

Dim LastX As Integer
Dim LastY As Integer

Dim LineColor
Dim YourName As String
Dim OtherName As String

Dim TransBuff As String
'In TransBuff we store the data (lines and colors) that were going to send to the other side
'the format i use is:
'lineX1,lineY1,lineX2,lineY2,lineColor¿
'if we draw this for each line into the transmit buffer and then send it to the other side
'and clear the buffer every 500ms (1/2 second), it doesnt get too big in size and its easy to decipher :)

Private Sub Command1_Click()
Picture1.Cls
ws.SendData "CLEAR" 'Clear the other sides picture too
End Sub

Private Sub Command2_Click()
cd.ShowColor
LineColor = cd.Color 'Set new color
End Sub

Private Sub Command3_Click() 'just code to save the image
cd.InitDir = App.Path
cd.Filter = "Bitmap Image (*.bmp)|*.bmp"
cd.DialogTitle = "Save WhiteBoardPro Drawing"
cd.ShowSave
If cd.FileName <> "" Then
SavePicture Picture1.Image, cd.FileName
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer) 'just sends chat when we press enter
If KeyAscii = 13 Then
Call Command4_Click
End If
End Sub

Private Sub Command4_Click()
ws.SendData "CHAT" & Text1.Text 'send a chat protocol, followed by the chat
CBox.Text = CBox.Text & YourName & ": " & Text1.Text & vbCrLf
CBox.SelStart = Len(CBox.Text)
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
LineColor = 0 'default linecolor = black
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmLogin
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'if user is clicking

'Draw it on our screen
Picture1.Line (LastX, LastY)-(X, Y), LineColor

'Copy the data into the transmit buffer in my format (see top)
TransBuff = TransBuff & LastX & "," & LastY & "," & X & "," & Y & "," & LineColor & "¿"

End If

LastX = X
LastY = Y

End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim WData As String
ws.GetData WData

If Left(WData, 3) = "DRW" Then 'Draw the other sides lines
Dim DrawBuff As String
DrawBuff = Right(WData, Len(WData) - 3) 'get rid of that "DRW"
Dim DLine
DLine = Split(DrawBuff, "¿") 'Split it into each line (see top for format)
    For i = 0 To (UBound(DLine) - 1)  'for i=0 to number of lines to be drawn
    tmp = Split(DLine(i), ",")
    Picture1.Line (tmp(0), tmp(1))-(tmp(2), tmp(3)), tmp(4) 'That draws the line into the pic box, colors n' all :)
    Next
End If

If InStr(1, WData, "CLEAR") Then Picture1.Cls 'Clear the picture box when told to

If Left(WData, 4) = "CHAT" Then 'incoming chat
CBox.Text = CBox.Text & OtherName & ": " & Right(WData, Len(WData) - 4) & vbCrLf
CBox.SelStart = Len(CBox.Text)
End If

If Left(WData, 4) = "NAME" Then 'incoming name
YourName = frmLogin.Text1.Text
OtherName = Right(WData, Len(WData) - 4)
CBox.Text = OtherName & " Enters" & vbCrLf
End If

End Sub

Private Sub TransTimer_Timer()
'DRW is just a protocol ive used
If Len(TransBuff) > 0 Then 'If we drawn something, send it

'The problem is, too much data crashes the program, cos winsock cant take it :(, so we need
'to make sure there isnt too much data - as it happens, im lazy and because ive never managed
'to send more than 1k in half a second, i just delete the data if it is more that 1k

If Len(TransBuff) > 1024 Then 'Thats 1kb (in bytes)
Exit Sub 'leave the sub, and dont send anything
End If

ws.SendData "DRW" & TransBuff 'Send the data to be drawn
TransBuff = "" 'Clear the buffer

End If
End Sub

Private Sub ws_Connect()

frmLogin.Hide
frmMain.Show
TransTimer.Enabled = True 'Start sending data
ws.SendData "NAME" & frmLogin.Text1.Text 'Send our name
frmLogin.TimeOut.Enabled = False 'Stop connection from timing out, cos we've conected!
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
ws.Close
ws.Accept requestID

frmLogin.Hide
frmMain.Show
TransTimer.Enabled = True 'Start sending data
ws.SendData "NAME" & frmLogin.Text1.Text 'Send our name
End Sub
