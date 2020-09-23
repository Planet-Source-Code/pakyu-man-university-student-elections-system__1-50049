VERSION 5.00
Begin VB.Form frmAdministrator 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Access"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4695
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "ADMINISTRATOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   6360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   7335
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton cmdParty 
         BackColor       =   &H00808080&
         Caption         =   "Party E&ntry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPositions 
         BackColor       =   &H00808080&
         Caption         =   "Candidate Position &Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00808080&
         Caption         =   "Delete All &Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton cmdRegStat 
         BackColor       =   &H00808080&
         Caption         =   "Course Entry / Voter &Statistics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCanvassVotes 
         BackColor       =   &H00808080&
         Caption         =   "Canvass &Votes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdminExit 
         BackColor       =   &H00808080&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCandidates 
         BackColor       =   &H00808080&
         Caption         =   "&Candidate Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.Image imgLogo 
      Height          =   3495
      Left            =   600
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3960
      TabIndex        =   8
      Top             =   7800
      Width           =   390
   End
End
Attribute VB_Name = "frmAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdminExit_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdCandidates_Click()
    frmCandidates.Show
    Unload Me
End Sub
Private Sub cmdCanvassVotes_Click()
    frmVoteResults.Show
    Unload Me
End Sub
Private Sub cmdClear_Click()
    ToPass = 4
    frmPassAdmin.Show
    Me.Enabled = False
End Sub
Private Sub cmdParty_Click()
    Unload Me
    frmParty.Show
End Sub
Private Sub cmdPositions_Click()
    Unload Me
    frmPositions.Show
End Sub

Private Sub cmdRegStat_Click()
    frmVoterStats.Show
    Unload Me
End Sub

Private Sub Form_Load()
'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    imgLogo.Picture = LoadPicture(MyPicture)
End Sub
Private Sub Timer1_Timer()
    lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
End Sub
