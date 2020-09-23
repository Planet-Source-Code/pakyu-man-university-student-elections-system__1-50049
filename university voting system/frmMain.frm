VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VOTING SYSTEM"
   ClientHeight    =   5040
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7335
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5295
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VOTING SYSTEM"
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
         TabIndex        =   7
         Top             =   240
         Width           =   2970
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   4440
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton cmdVoterRegistration 
         BackColor       =   &H00808080&
         Caption         =   "V&oter Registration"
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
      Begin VB.CommandButton cmdElection 
         BackColor       =   &H00808080&
         Caption         =   "&Election Proper"
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
      Begin VB.CommandButton cmdMainExit 
         BackColor       =   &H00808080&
         Caption         =   "E&xit System"
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
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdminAccess 
         BackColor       =   &H00808080&
         Caption         =   "&Administrator"
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
   End
   Begin VB.Image imgLogo 
      Height          =   3255
      Left            =   960
      Stretch         =   -1  'True
      ToolTipText     =   "Double-click here to paste picture."
      Top             =   1080
      Width           =   3495
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
      Left            =   4440
      TabIndex        =   5
      Top             =   4560
      Width           =   390
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Info"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MySys As FileSystemObject

Private Sub cmdCandi_Click()
    frmCandidates.Show
End Sub

Private Sub cmdAdminAccess_Click()
    ToPass = 1
    frmPassAdmin.Show
    Me.Enabled = False
End Sub
Private Sub cmdElection_Click()
    frmVote.Show
    Unload Me
End Sub
Private Sub cmdMainExit_Click()
    Unload frmBack
    End
End Sub

Private Sub cmdVoterRegistration_Click()
    frmVoterRegistration.Show
    Unload Me
End Sub
Private Sub Form_DblClick()
    imgLogo.Picture = LoadPicture("")
On Error GoTo 100
    Kill "c:\esnSystem.dat"
100
End Sub

Private Sub Form_Load()
    ToPass = 0
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    frmBack.Show
    frmBack.Enabled = False
    Call PicLoader
    
On Error GoTo Pro
    MkDir App.Path + "\VOTING SYSTEM"
Pro:
End Sub

Private Sub imgLogo_DblClick()
    
        On Error GoTo Handler
        CommonDialog1.CancelError = True
   
        CommonDialog1.FileName = ""
        CommonDialog1.Filter = "all files (*.*)|*.*|bmp files (*.bmp)|*.bmp|jpeg files (*.jpg)|*.jpg|ico files (*.ico)|*.ico|gif files (*.gif)|*.gif"
        CommonDialog1.ShowOpen
        MyPicture = CommonDialog1.FileName
    
        Open "c:\esnSystem.dat" For Output As #505
        Write #505, MyPicture
        Close #505
    
        imgLogo.Picture = LoadPicture(MyPicture)
Handler:
End Sub

Private Sub mnuAbout_Click()
    frmAboutVoteSys.Show
    Me.Enabled = False
End Sub

Private Sub Timer1_Timer()
     lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
End Sub

Private Sub PicLoader()
On Error GoTo 100
    Open "c:\esnSystem.dat" For Input As #500
    Input #500, MyPicture
    Close #500
    imgLogo.Picture = LoadPicture(MyPicture)
    Exit Sub
100
    Open "c:\esnSystem.dat" For Output As #501
    Write #501, ""
    Close #501
    
    Open "c:\esnSystem.dat" For Input As #502
    Input #502, MyPicture
    Close #502
    
imgLogo.Picture = LoadPicture(MyPicture)
End Sub
