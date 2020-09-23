VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "SHOUTcast Streaming Audio Capture"
   ClientHeight    =   3990
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   0
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4560
      Top             =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   2415
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1200
      Width           =   5415
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   0
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Stream"
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Stream"
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lS 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   3720
      Width           =   45
   End
   Begin VB.Label lState 
      AutoSize        =   -1  'True
      Caption         =   "State:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save as.."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   195
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu FileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bIcy As Boolean
Public ff As Integer
Public lIP, lPort As Long
Public headerLoc As Long
Const reqHeader = "GET / HTTP/1.0" & vbLf & vbLf
Const endHeader = vbCr & vbLf & vbCr & vbLf


Private Sub Send(ByVal lpBuf As String, ByVal nBufLen As Integer, Optional nFlags As Integer = 0)
ff = FreeFile
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then Exit Sub

Command1.Enabled = False
Command3.Enabled = True

WS.Connect CStr(Text1.Text), CStr(Text2.Text)

Do While WS.State <> sckConnected
    If WS.State = sckError Then
        MsgBox "Err"
        Exit Sub
    End If
    DoEvents
Loop

Open Text3.Text For Binary As ff

WS.SendData reqHeader
End Sub

Private Sub CloseIP()
On Error Resume Next
Close #ff
WS.Close
bIcy = True
Text4.Text = ""
End Sub

Private Sub Command1_Click()

Call Send(reqHeader, Len(reqHeader))
End Sub

Private Sub Command2_Click()
CloseIP
Unload Me
End Sub

Private Sub Command3_Click()
Command1.Enabled = True
Command3.Enabled = False
CloseIP

End Sub

Private Sub Command4_Click()
CMD1.DialogTitle = "Create File In..."
CMD1.ShowSave
Text3.Text = CMD1.FileName & ".mp3"
End Sub

Private Sub FileExit_Click()
CloseIP
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = WS.LocalIP
Text2.Text = "8000"
Text3.Text = "C:\Windows\Desktop\Stream.mp3"
Command3.Enabled = False
bIcy = True
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Close #ff
WS.Close
End Sub

Private Sub HelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub Timer1_Timer()
Dim a As StateConstants
Select Case WS.State
Case sckClosed
    lS.Caption = "Socket Closed"
Case sckClosing
    lS.Caption = "Socket Closing..."
Case sckConnected
    lS.Caption = "Socket Connected"
Case sckConnecting
    lS.Caption = "Socket Connecting"
Case sckConnectionPending
    lS.Caption = "Socket Connection Pending..."
Case sckError
    lS.Caption = "Socket Encountered Error"
Case sckHostResolved
    lS.Caption = "Socket Resolved Host"
Case sckListening
    lS.Caption = "Socket is Listening"
Case sckOpen
    lS.Caption = "Socket is Open"
Case sckResolvingHost
    lS.Caption = "Socket Resolving Host..."
End Select
End Sub

Private Sub WS_ConnectionRequest(ByVal requestID As Long)
WS.Accept requestID
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim ret As Integer
Dim inBuf As String * 1024
WS.GetData inBuf, , 1024

If bIcy = True Then 'Still need to parse out Header
    For i = 1 To Len(inBuf)
        If Mid(inBuf, i, 4) = endHeader Then
            Text4.Text = Text4.Text & Left(inBuf, i + 4)
            Put #ff, Loc(ff) + 1, Right(inBuf, Len(inBuf) + 4)
            headerLoc = i
            bIcy = False
            'MsgBox "header Found at loc " & i
        End If
    Next i
Else
    Put #ff, Loc(ff) + 1, inBuf
End If
End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
End Sub
