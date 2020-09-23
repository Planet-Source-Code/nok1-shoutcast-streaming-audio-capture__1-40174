VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SHOUTcast Streaming Audio Capture"
   ClientHeight    =   3090
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":0000
   ScaleHeight     =   2132.773
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4440
      TabIndex        =   0
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "http://www.shoutcast.com/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      MouseIcon       =   "frmAbout.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2160
      Width           =   1980
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4440
      MouseIcon       =   "frmAbout.frx":0B8E
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Additional info on SHOUTcast can be found at"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   3300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nok1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2640
      MouseIcon       =   "frmAbout.frx":0FD0
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1200
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Visual Basic Code by:"
      Height          =   195
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chris Hartmann"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2280
      MouseIcon       =   "frmAbout.frx":1412
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Original Code by:"
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5224.884
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":1854
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   " Shoutcast Streaming Audio Capture"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5210.798
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0 VB6"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   720
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Label2_Click()
    ShellExec "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=2347&lngWId=3", "open", "", False
End Sub

Private Sub Label4_Click()
    ShellExec "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&txtCriteria=Nok1&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&B1=Quick+Search", "open", "", False
End Sub

Private Sub Label7_Click()
ShellExec "http://www.shoutcast.com/", "open", "", False
End Sub
