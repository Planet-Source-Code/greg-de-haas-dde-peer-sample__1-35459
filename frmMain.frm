VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DDE Peer"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkMode        =   1  'Source
   LinkTopic       =   "Peer"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   330
      Left            =   -45
      ScaleHeight     =   270
      ScaleWidth      =   4725
      TabIndex        =   4
      Top             =   5400
      Width           =   4785
      Begin VB.Label lblStatus 
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   900
         TabIndex        =   6
         Top             =   45
         Width           =   3570
      End
      Begin VB.Label Label3 
         Caption         =   "Status: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.Timer tmrConnect 
      Interval        =   500
      Left            =   4005
      Top             =   0
   End
   Begin VB.TextBox txtSend 
      Height          =   2310
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2970
      Width           =   4245
   End
   Begin VB.TextBox txtReceive 
      Height          =   2310
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   315
      Width           =   4245
   End
   Begin VB.Label Label2 
      Caption         =   "Text To Send   (Type Here)"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   2745
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Incoming Text"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Timer to attempt to connect
Private Sub tmrConnect_Timer()
On Error GoTo DoError
    txtReceive.LinkMode = 0
    txtReceive.LinkTopic = "prjDDE|Peer"
    txtReceive.LinkItem = "txtSend"
    txtReceive.LinkTimeout = 100
    txtReceive.LinkMode = vbLinkAutomatic
    
    'At this point, we have a successful connection
    'tmrConnect.Enabled = False  'Connected
    lblStatus = "Connected!"
    lblStatus.ForeColor = RGB(0, 128, 0)
    
    GoTo DoEnd
DoError:
    'MsgBox "DDE Link Failed : " & Err.Description
    On Error Resume Next
    txtReceive.LinkMode = 0
    lblStatus = "No Connection - Peer Not Found"
    lblStatus.ForeColor = RGB(128, 0, 0)
    Resume DoEnd
DoEnd:
End Sub
