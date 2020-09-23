VERSION 5.00
Begin VB.Form frmVote 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VOTER REGISTRATION"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmeVoteDefaultScreen 
      BackColor       =   &H00000000&
      Height          =   7095
      Left            =   960
      TabIndex        =   16
      Top             =   360
      Width           =   9855
      Begin VB.Frame Frame7 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   9375
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "ELECTION"
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
            TabIndex        =   19
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Height          =   2295
         Left            =   7680
         TabIndex        =   17
         Top             =   4440
         Width           =   1935
         Begin VB.Frame Frame4 
            BackColor       =   &H00000000&
            Height          =   1335
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   1695
            Begin VB.CommandButton cmdEnter 
               BackColor       =   &H00808080&
               Caption         =   "&Enter and Vote"
               Default         =   -1  'True
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
               TabIndex        =   23
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.TextBox txtStudNumReg 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            Caption         =   "STUDENT NUMBER:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1470
         End
      End
      Begin VB.Image imgLogo 
         Height          =   3495
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   7920
         TabIndex        =   20
         Top             =   4560
         Width           =   45
      End
   End
   Begin VB.Frame fmeFinal 
      BackColor       =   &H80000007&
      Height          =   6855
      Left            =   3240
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdVote 
         Appearance      =   0  'Flat
         Caption         =   "&Vote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelVote 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel (Repeat Entries)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6000
         Width           =   1455
      End
      Begin VB.TextBox txtFinalVotes 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label lblVoterName 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "YOUR CURRENT VOTE ENTRIES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3630
      End
   End
   Begin VB.Frame fmeStart 
      BackColor       =   &H00000000&
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame9 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   975
         Left            =   120
         TabIndex        =   30
         Top             =   6240
         Width           =   6855
         Begin VB.CommandButton cmdChoose 
            Appearance      =   0  'Flat
            Caption         =   "&Choose"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel My &Previous Vote Entries"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4560
            TabIndex        =   32
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmdNoVote 
            Appearance      =   0  'Flat
            Caption         =   "&No Vote For This Position"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         Height          =   4695
         Left            =   7080
         TabIndex        =   27
         Top             =   2520
         Width           =   4455
         Begin VB.TextBox txtVotes 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   28
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Current Vote Entries:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   2250
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000007&
         Caption         =   "Photo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3735
         Left            =   3240
         TabIndex        =   26
         Top             =   2520
         Width           =   3735
         Begin VB.Image imgPic 
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Left            =   120
            Stretch         =   -1  'True
            ToolTipText     =   "Double-click to set photo."
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000007&
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   10935
         Begin VB.Label lblPos 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "ELECTION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   2100
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "ELECTION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000007&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   10935
         Begin VB.Label lblParty 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "ELECTION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1470
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00000000&
         Caption         =   "Candidates"
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
         Height          =   3735
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   3015
         Begin VB.ListBox lstCandidates 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3210
            ItemData        =   "frmVote.frx":0000
            Left            =   120
            List            =   "frmVote.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10935
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "ELECTION"
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
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   6960
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
      Left            =   7920
      TabIndex        =   1
      Top             =   7560
      Width           =   390
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gPerson As Positions
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim gCurrentRecord As Long
Dim gLastRecord As Long

Dim gPersonX As Candidates
Dim gFileNumX As Integer
Dim gRecordLenX As Long
Dim gCurrentRecordX As Long
Dim gLastRecordX As Long

Dim MySys As New FileSystemObject
Dim NewRec As Boolean
Dim NewFile As Boolean
Dim NewRecX As Boolean
Dim NewFileX As Boolean
Dim hold As String

Dim PosHold As String
Dim PartyHold As String



Private Sub cmdCancel_Click()
    txtVotes.Text = ""
    Close #gFileNum
    Close #gFileNumX
    StartVote
    lstCandidates.SetFocus
End Sub

Private Sub cmdCancelVote_Click()
    txtFinalVotes.Text = ""
    txtVotes.Text = ""
    fmeFinal.Visible = False
    fmeStart.Visible = True
    cmdChoose.Default = True
    Close #gFileNum
    Close #gFileNumX
    StartVote
    lstCandidates.SetFocus
End Sub

Private Sub cmdChoose_Click()
    If lstCandidates = "" Then
        Call cmdNoVote_Click
    Else
        txtVotes.Text = txtVotes.Text + vbNewLine + " FOR " + lblPos.Caption + ": " + lblName.Caption + vbNewLine + " (" + lblParty.Caption + ")" + vbNewLine
        InitializeVote
        gCurrentRecord = gCurrentRecord + 1
        If gCurrentRecord > gLastRecord Then
            fmeFinal.Visible = True
            fmeStart.Visible = False
            txtFinalVotes.Text = txtVotes.Text
            txtVotes.Text = ""
            cmdVote.Default = True
            cmdVote.SetFocus
            Exit Sub
        End If
        ShowCurrentPosition
        ListOfCandi
        lstCandidates.SetFocus
    End If
End Sub

Private Sub cmdEnter_Click()
'to open VOTERS
    
    If txtStudNumReg.Text = "" Then
        MsgBox "Enter your student number to vote!", vbOKOnly + vbExclamation, "Vote"
        txtStudNumReg.SetFocus
        Exit Sub
    End If
    
    Dim gPersonB As VoterInfo
    Dim gFileNumB As Integer
    Dim gRecordLenB As Long
    Dim gCurrentRecordB As Long
    Dim gLastRecordB As Long
    
    gRecordLenB = Len(gPersonB)

    gFileNumB = FreeFile

    Open App.Path + "\VOTING SYSTEM\Voters.dat" For Random As gFileNumB Len = gRecordLenB

    gCurrentRecordB = 1
    gLastRecordB = FileLen(App.Path + "\VOTING SYSTEM\Voters.dat") / gRecordLenB
    
    If gLastRecordB = 0 Then
        gLastRecordB = 1
    End If
    
        
    Dim RecNumX As Integer
    Dim NameToSearch As String
    
    NameToSearch = UCase(txtStudNumReg.Text)
    
    For RecNumX = 1 To gLastRecordB
        Get #gFileNumB, RecNumX, gPersonB
        If NameToSearch = Trim(gPersonB.VStudNum) And Trim(gPersonB.VVoteTrack) = 0 Then
            lblVoterName.Caption = Trim(gPersonB.VFullName)
            fmeVoteDefaultScreen.Visible = False
            fmeStart.Visible = True
            gPersonB.VVoteTrack = 1
            Put #gFileNumB, RecNumX, gPersonB
            MsgBox "Welcome " + UCase(Trim(gPersonB.VFirstName)) + ".", vbOKOnly + vbInformation, "Vote"
            txtStudNumReg.Text = ""
            Close #gFileNumB
            Call StartVote
            cmdChoose.Default = True
            Exit Sub
        ElseIf NameToSearch = Trim(gPersonB.VStudNum) And Trim(gPersonB.VVoteTrack) = 1 Then
            MsgBox "You already have voted " + UCase(Trim(gPersonB.VFirstName)) + ". You cannot repeat votes.", vbOKOnly + vbExclamation, "Vote"
            txtStudNumReg.Text = ""
            txtStudNumReg.SetFocus
            Close #gFileNumB
            Exit Sub
       
        End If
    Next
    
    MsgBox "You are not a registered voter.  See System Administrator for registration.", vbOKOnly + vbExclamation, "Vote"
    txtStudNumReg.Text = ""
    txtStudNumReg.SetFocus
    Close #gFileNumB
    
End Sub

Private Sub cmdEnter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        ToPass = 2
        frmPassAdmin.Show
        Me.Enabled = False
    End If
  
End Sub

Private Sub cmdNoVote_Click()
    txtVotes.Text = txtVotes.Text + vbNewLine + " FOR " + lblPos.Caption + ": NO VOTE " + vbNewLine
    If lstCandidates = "" Then
        txtVotes.Text = txtVotes.Text + " (NO AVAILABLE CANDIDATES TO VOTE)" + vbNewLine
    End If
    gCurrentRecord = gCurrentRecord + 1
    If gCurrentRecord > gLastRecord Then
        fmeFinal.Visible = True
        fmeStart.Visible = False
        txtFinalVotes.Text = txtVotes.Text
        Exit Sub
    End If
    ShowCurrentPosition
    ListOfCandi
    lstCandidates.SetFocus
    
    
End Sub

Private Sub cmdVote_Click()
    AddVotes
  
End Sub


Private Sub Form_Load()
'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

'Call StartVote
    imgLogo.Picture = LoadPicture(MyPicture)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Close #gFileNumX
    Close #gFileNum
End Sub



Private Sub lstCandidates_Click()
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As Candidates
    Dim RecNumX As Integer
    
    NameToSearch = lstCandidates
    If NameToSearch = "" Then
        Exit Sub
    End If
    
    NameToSearch = UCase(NameToSearch)
    Found = False
    For RecNum = 1 To gLastRecordX
        Get #gFileNumX, RecNum, TmpPerson
        If NameToSearch = UCase(Trim(TmpPerson.CFullName)) Then
            Found = True
            Exit For
        End If
    Next
    If Found = True Then
        gCurrentRecordX = RecNum
        Get #gFileNumX, gCurrentRecordX, gPersonX
        lblParty.Caption = Trim(gPersonX.CParty)
        lblName.Caption = Trim(gPersonX.CFullName)
        On Error GoTo ErrorOnPic
        imgPic.Picture = LoadPicture(Trim(gPersonX.CPic))
        Exit Sub
    
ErrorOnPic:
        imgPic.Picture = LoadPicture("")
    Else
        MsgBox "Candidate " + NameToSearch + " not found.", vbOKOnly + vbExclamation, "Search"
    End If
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
End Sub
Private Sub StartVote()


'to open the CANDIDATE DATABASE
    gRecordLen = Len(gPerson)

    gFileNum = FreeFile

    Open App.Path + "\VOTING SYSTEM\Positions.dat" For Random As gFileNum Len = gRecordLen

    gCurrentRecord = 1
    gLastRecord = FileLen(App.Path + "\VOTING SYSTEM\Positions.dat") / gRecordLen
  
    If gLastRecord = 0 Then
        gLastRecord = 1
        NewFile = True
    End If
    
'to open CANDIDATES
    
    gRecordLenX = Len(gPersonX)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNumX Len = gRecordLenX

    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Candi.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    
    Call ShowCurrentPosition
    Call ListOfCandi
   
        
    
End Sub

Private Sub ShowCurrentPosition()
    Get #gFileNum, gCurrentRecord, gPerson
    
    
    lblPos.Caption = Trim(gPerson.Position)
                 
End Sub

Private Sub ListOfCandi()
    
    lstCandidates.Clear
    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
        If Trim(gPersonX.CPos) = lblPos.Caption Then
            lstCandidates.AddItem Trim(gPersonX.CFullName)
                hold = Trim(gPersonX.CFullName)
        End If
    Next
    lblName.Caption = hold
    lstCandidates = lblName.Caption
    
  
End Sub

Private Sub InitializeVote()
    Dim RecNumX As Integer

    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
        If Trim(gPersonX.CFullName) = lblName.Caption Then
        gPersonX.CAddVote = 1
    
        Put #gFileNumX, RecNumX, gPersonX
        
        End If
    Next
End Sub

Private Sub AddVotes()
    
    Dim RecNumX As Integer
    Dim Response As String
    Response = MsgBox("Are you sure of your votes?", vbYesNo + vbQuestion, "Votes")
    
    If Response = vbYes Then
        For RecNumX = 1 To gLastRecordX
            Get #gFileNumX, RecNumX, gPersonX
            If Trim(gPersonX.CAddVote) = 1 Then
                gPersonX.CAddVote = 1
                gPersonX.CVote = gPersonX.CVote + 1
                gPersonX.CAddVote = 0
                Put #gFileNumX, RecNumX, gPersonX
            End If
        Next
        PlaceVotes
        MsgBox "Your vote has been submitted. Thank you.", vbOKOnly + vbExclamation, "Election"
        fmeFinal.Visible = False
        fmeVoteDefaultScreen.Visible = True
        cmdEnter.Default = True
        Close #gFileNum
        Close #gFileNumX
        txtVotes.Text = ""
    Else
        cmdVote.SetFocus
    End If
    
End Sub

Private Sub PlaceVotes()
    
    Dim gPersonB As VoterInfo
    Dim gFileNumB As Integer
    Dim gRecordLenB As Long
    Dim gCurrentRecordB As Long
    Dim gLastRecordB As Long
    
    gRecordLenB = Len(gPersonB)

    gFileNumB = FreeFile

    Open App.Path + "\VOTING SYSTEM\Voters.dat" For Random As gFileNumB Len = gRecordLenB

    gCurrentRecordB = 1
    gLastRecordB = FileLen(App.Path + "\VOTING SYSTEM\Voters.dat") / gRecordLenB
    
    If gLastRecordB = 0 Then
        gLastRecordB = 1
    End If
    
        
    Dim RecNumX As Integer
    Dim NameToSearch As String
    
    NameToSearch = UCase(lblVoterName.Caption)
    
    For RecNumX = 1 To gLastRecordB
        Get #gFileNumB, RecNumX, gPersonB
        If NameToSearch = Trim(gPersonB.VFullName) Then
            gPersonB.VVote = txtFinalVotes.Text
            gPersonB.VVoteTimeTrack = lblTime.Caption
            Put #gFileNumB, RecNumX, gPersonB
            Close #gFileNumB
            Exit Sub
        End If
    Next
    
    Close #gFileNumB
End Sub
