VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCandidates 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CANDIDATES"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fmeStart 
      BackColor       =   &H80000007&
      Height          =   5055
      Left            =   120
      TabIndex        =   33
      Top             =   360
      Width           =   11655
      Begin VB.Frame Frame7 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   11415
         Begin VB.Label lblParty 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Height          =   2415
         Left            =   9840
         TabIndex        =   35
         Top             =   2160
         Width           =   1695
         Begin VB.CommandButton cmdEnter 
            BackColor       =   &H00808080&
            Caption         =   "&Enter"
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
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdVoteExit 
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
            TabIndex        =   2
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   2415
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   9615
         Begin VB.ListBox lstParty 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            ItemData        =   "frmCandidates.frx":0000
            Left            =   120
            List            =   "frmCandidates.frx":0002
            TabIndex        =   0
            Top             =   240
            Width           =   9375
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Choose party for candidate entry then press enter."
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
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   4215
      End
   End
   Begin VB.Frame fmeEntry 
      BackColor       =   &H80000007&
      Height          =   5415
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame1 
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
         Height          =   2295
         Left            =   7320
         TabIndex        =   41
         Top             =   120
         Width           =   2415
         Begin VB.Image imgPic 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Left            =   240
            Stretch         =   -1  'True
            ToolTipText     =   "Doublie - click to set picture."
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtHolder 
         Height          =   285
         Left            =   3480
         TabIndex        =   39
         Top             =   4800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtCStudNo 
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
         TabIndex        =   10
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboCCourse 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   3360
         Width           =   4695
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00000000&
         Caption         =   "Running Positions"
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
         Height          =   2775
         Left            =   7320
         TabIndex        =   27
         Top             =   2520
         Width           =   2415
         Begin VB.ListBox lstPos 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2400
            ItemData        =   "frmCandidates.frx":0004
            Left            =   120
            List            =   "frmCandidates.frx":0006
            TabIndex        =   3
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.TextBox txtCMidName 
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
         Left            =   4920
         TabIndex        =   14
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox txtCFirstName 
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
         Left            =   2520
         TabIndex        =   13
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox txtCLastName 
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
         TabIndex        =   12
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   7095
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Caption         =   "FOR "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   465
         End
         Begin VB.Label lblPosition 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   600
            TabIndex        =   26
            Top             =   240
            Width           =   45
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   7095
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "CANDIDATE ENTRY"
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
            TabIndex        =   24
            Top             =   240
            Width           =   3405
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000012&
         Height          =   5175
         Left            =   9840
         TabIndex        =   22
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton cmdPicture 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Set &Photo"
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
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmCandidates.frx":0008
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdNew 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&New"
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
            Left            =   120
            Picture         =   "frmCandidates.frx":044A
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Input New Record"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Delete &All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaskColor       =   &H8000000A&
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Edit"
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
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmCandidates.frx":0AB4
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdCandiSave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Save"
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
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmCandidates.frx":0EF6
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Delete"
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
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmCandidates.frx":1560
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton cmdCandiExit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Back"
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
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmCandidates.frx":1662
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   4320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   7095
         Begin VB.Label lblPartyDisplay 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   7095
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.Label lblPicHolder 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   4920
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "LastName"
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
         Left            =   120
         TabIndex        =   32
         Top             =   3960
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Student Number"
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
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Course"
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
         Left            =   2520
         TabIndex        =   30
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Middlename"
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
         Left            =   4920
         TabIndex        =   29
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Firstname"
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
         Left            =   2520
         TabIndex        =   28
         Top             =   3960
         Width           =   840
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   5160
   End
   Begin VB.Label lblRecNum 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   45
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
      Left            =   8760
      TabIndex        =   15
      Top             =   5640
      Width           =   390
   End
End
Attribute VB_Name = "frmCandidates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gPerson As Candidates
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim gCurrentRecord As Long
Dim gLastRecord As Long
Dim MySys As New FileSystemObject
Dim NewRec As Boolean
Dim NewFile As Boolean
Dim PosHold As String
Dim PartyHold As String
Private Sub cboCCourse_GotFocus()
    cboCCourse.SelStart = 0
    cboCCourse.SelLength = Len(cboCCourse.Text)
End Sub

Private Sub cmdCandiExit_Click()
    ChangeCatcher
    fmeEntry.Visible = False
    fmeStart.Visible = True
    cmdEnter.Default = True
    
End Sub

Private Sub cmdCandiSave_Click()
    If txtCStudNo.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtCStudNo.SetFocus
        Exit Sub
    End If
    If cboCCourse.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        cboCCourse.SetFocus
        Exit Sub
    End If
    If txtCLastName.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtCLastName.SetFocus
        Exit Sub
    End If
    If txtCFirstName.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtCFirstName.SetFocus
        Exit Sub
    End If
    If txtCMidName.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtCMidName.SetFocus
        Exit Sub
    End If
        
    If NewRec = True And NewFile = False Then
            gLastRecord = gLastRecord + 1
            Put #gFileNum, gLastRecord, gPerson
            gCurrentRecord = gLastRecord
            SaveCandi
            cmdNew.Enabled = False
            MsgBox "Candidate saved.", vbOKOnly + vbInformation, "Save"
            ShowCandi
    Else
            SaveCandi
            cmdNew.Enabled = False
            MsgBox "Candidate saved.", vbOKOnly + vbInformation, "Save"
            ShowCandi
            NewFile = False
    End If
    On Error GoTo 100
    lblPosition.Caption = lstPos
    lstPos.SetFocus
    Exit Sub
100
     lblPosition.Caption = lstPos
    lstPos.SetFocus
End Sub

Private Sub cmdClear_Click()
    Msg = "Are you sure you want ot delete all records?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbYes Then
        If MySys.FileExists(App.Path + "\VOTING SYSTEM\Candi.dat") = True Then
            Close #gFileNum
            Kill App.Path + "\VOTING SYSTEM\Candi.dat"
            
            gFileNum = FreeFile
            Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNum Len = gRecordLen
    
            gLastRecord = 1
            gCurrentRecord = gLastRecord
            NewFile = True
            lblRecNum.Visible = False
            NewRec = False
            MsgBox "All records cleared!", vb0konly + vbInformation, "Delete"
            Call lstPos_Click
        End If
    End If
    lblPosition.Caption = lstPos
    lstPos.SetFocus
    NewRec = True
    NewFile = True
            
End Sub

Private Sub cmdDelete_Click()
'to delete records
    If lblName.Caption = UCase("no entry yet") Then
        MsgBox "There is no record selected to delete.", vbOKOnly + vbExclamation, "Delete"
        lblPosition.Caption = lstPos
        lstPos.SetFocus
        Exit Sub
    End If
    
    Dim DirResult
    Dim TmpFileNum
    Dim TmpPerson As Candidates
    Dim RecNum As Long
    Dim TmpRecNum As Long
    Dim Msg As String
    Msg = "Delete " + UCase(txtCFirstName.Text + " " + txtCLastName.Text + "'s record?")
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbNo Then
        lblPosition.Caption = lstPos
        lstPos.SetFocus
        Exit Sub
    End If
    
    
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\MyCandi.tmp") = True Then
        Kill App.Path + "\VOTING SYSTEM\MyCandi.tmp"
    End If
    TmpFileNum = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\MyCandi.tmp" For Random As TmpFileNum Len = gRecordLen
        
    RecNum = 1
    TmpRecNum = 1
    Do While RecNum < gLastRecord + 1
        If RecNum <> gCurrentRecord Then
            Get #gFileNum, RecNum, TmpPerson
            Put #TmpFileNum, TmpRecNum, TmpPerson
            TmpRecNum = TmpRecNum + 1
        End If
        RecNum = RecNum + 1
    Loop
    
    Close #gFileNum
    
    MySys.DeleteFile App.Path + "\VOTING SYSTEM\Candi.dat"
    
    Close #TmpFileNum
    Name App.Path + "\VOTING SYSTEM\MyCandi.tmp" As App.Path + "\VOTING SYSTEM\Candi.dat"
    
    gFileNum = FreeFile
    Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNum Len = gRecordLen
    
    gLastRecord = gLastRecord - 1
    If gLastRecord = 0 Then
        gLastRecord = 1
        NewFile = True
    End If
    
    If gCurrentRecord > gLastRecord Then
        gCurrentRecord = gLastRecord
    End If
    
    MsgBox "Candidate deleted.", vb0konly + vbInformation, "Delete"
    Call lstPos_Click
    
    lblPosition.Caption = lstPos
    lstPos.SetFocus
    NewRec = True
End Sub

Private Sub cmdEdit_Click()
    txtCStudNo.SetFocus
End Sub

Private Sub cmdEnter_Click()

    fmeStart.Visible = False
    fmeEntry.Visible = True
    Call ListOfPositions
 
    lblPartyDisplay.Caption = lblParty.Caption
    lstPos.SetFocus
    
    lstPos = PosHold
    cmdCandiSave.Default = True
    If lstPos = "" Then
        MsgBox "You cannot enter candidates without entering positions first.  Exit then go to Position Entry first.", vbOKOnly + vbExclamation, "Error"
        Call cmdCandiExit_Click
    End If
    Call ListOfCourse
End Sub

Private Sub cmdNew_Click()
    Call ChangeCatcher
    
    NewRec = True
    lblRecNum.Visible = False
    
    gPerson.CAddVote = 0
    gPerson.CVote = 0
    gPerson.CCourse = ""
    gPerson.CFirstName = ""
    gPerson.CFullName = ""
    gPerson.CLastName = ""
    gPerson.CMidName = ""
    gPerson.CParty = ""
    gPerson.CPos = ""
    gPerson.CStudNum = ""
    gPerson.CPic = ""
    
    cboCCourse.Text = ""
    txtCLastName.Text = ""
    txtCMidName.Text = ""
    txtCFirstName.Text = ""
    txtCStudNo.Text = ""
    lblPicHolder.Caption = ""
    imgPic.Picture = LoadPicture(lblPicHolder.Caption)
    
    lblName.Caption = "NO ENTRY YET"
    
    cboCCourse.Tag = ""
    lblPicHolder.Tag = ""
    txtCLastName.Tag = ""
    txtCMidName.Tag = ""
    txtCFirstName.Tag = ""
    txtCStudNo.Tag = ""
    
    lblPosition.Caption = lstPos
    
    txtCStudNo.SetFocus
    
End Sub

Private Sub cmdPicture_Click()
    Call imgPic_DblClick
End Sub

Private Sub cmdVoteExit_Click()
    frmAdministrator.Show
    Unload Me
End Sub



Private Sub Form_Load()
'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

'to determine if there are no currently saved candidates
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\Candi.dat") = True Then
        NewFile = False
    Else
        NewFile = True
        MsgBox "No candidates are currently saved!", vbOKOnly + vbInformation, "Candidates"
    End If
   

'to open the CANDIDATES DATABASE
    gRecordLen = Len(gPerson)

    gFileNum = FreeFile

    Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNum Len = gRecordLen

    gCurrentRecord = 1
    gLastRecord = FileLen(App.Path + "\VOTING SYSTEM\Candi.dat") / gRecordLen
  
    If gLastRecord = 0 Then
        gLastRecord = 1
        NewFile = True
    End If
    Call ListOfParty
    lstParty = PartyHold
    If lstParty = "" Then
        MsgBox "You cannot enter candidates without entering parties first.  Exit then go to Party Entry first.", vbOKOnly + vbExclamation, "Error"
        cmdEnter.Enabled = False
        lstParty.Enabled = False
    End If
End Sub

Private Sub ShowCandi()
'to display the candidates
    Get #gFileNum, gCurrentRecord, gPerson
    
    txtCLastName.Text = Trim(gPerson.CLastName)
    txtCFirstName.Text = Trim(gPerson.CFirstName)
    txtCMidName.Text = Trim(gPerson.CMidName)
    txtCStudNo.Text = Trim(gPerson.CStudNum)
    cboCCourse.Text = Trim(gPerson.CCourse)
    lblPicHolder.Caption = Trim(gPerson.CPic)
    
    lblName.Caption = Trim(gPerson.CFullName)
    lblPosition.Caption = lstPos
    
    
    txtCLastName.Tag = txtCLastName.Text
    txtCFirstName.Tag = txtCFirstName.Text
    txtCMidName.Tag = txtCMidName.Text
    txtCStudNo.Tag = txtCStudNo.Text
    cboCCourse.Tag = cboCCourse.Text
    lblPicHolder.Tag = lblPicHolder.Caption
    
On Error GoTo ErrorOnPic
    imgPic.Picture = LoadPicture(lblPicHolder.Caption)
    
    lblRecNum.Visible = True
    lblRecNum.Caption = "Candidate #" + _
                    Str(gCurrentRecord) + " of " + _
                    Str(gLastRecord) + " saved candidate(s)."
    NewRec = False
    Exit Sub
ErrorOnPic:
    MsgBox "Error loading picture.  Photo will be saved blank.", vbOKOnly + vbExclamation, "Photo Error"
    lblPicHolder.Caption = ""
    lblPicHolder.Tag = lblPicHolder.Caption
    SaveCandi
    ShowCandi
    lblRecNum.Visible = True
    lblRecNum.Caption = "Candidate #" + _
                    Str(gCurrentRecord) + " of " + _
                    Str(gLastRecord) + " saved candidate(s)."
    NewRec = False
    
End Sub

Private Sub SaveCandi()
'to save candidates
    
    gPerson.CLastName = txtCLastName.Text
    gPerson.CFirstName = txtCFirstName.Text
    gPerson.CMidName = txtCMidName.Text
    gPerson.CStudNum = txtCStudNo.Text
    gPerson.CCourse = cboCCourse.Text
    gPerson.CParty = lblPartyDisplay.Caption
    gPerson.CPos = lblPosition.Caption
    gPerson.CFullName = UCase(txtCLastName.Text + ", " + txtCFirstName.Text + " " + Left(txtCMidName.Text, 1) + ".")
    gPerson.CAddVote = 0
    gPerson.CPic = lblPicHolder.Caption
    
    Put #gFileNum, gCurrentRecord, gPerson
    
    NewRec = False
    
    
    
End Sub

Private Sub ChangeCatcher()
    
    If txtCLastName.Tag <> txtCLastName.Text Then
        GoTo SaveNow
    End If
    If lblPicHolder.Tag <> lblPicHolder.Caption Then
        GoTo SaveNow
    End If
    If txtCFirstName.Tag <> txtCFirstName.Text Then
        GoTo SaveNow
    End If
    If txtCMidName.Tag <> txtCMidName.Text Then
        GoTo SaveNow
    End If
    If txtCStudNo.Tag <> txtCStudNo.Text Then
        GoTo SaveNow
    End If
    Exit Sub
SaveNow:
        Get #gFileNum, gCurrentRecord, gPerson
        If lblName.Caption = "" Then
            Msg = "Save changes to this new record?"
        Else
            Msg = "Save changes to the record of " + UCase(lblName.Caption) + "?"
        End If
        Response = MsgBox(Msg, vbYesNo + vbQuestion, "Save" & Save)
        If Response = vbYes Then
            
            Call cmdCandiSave_Click
        Else
            Exit Sub
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAdministrator.Show
    
    Close #gFileNum
End Sub

Private Sub imgPic_DblClick()
    If lblPicHolder.Caption <> "" Then
        imgPic.Picture = LoadPicture("")
        lblPicHolder.Caption = ""
    Else
        On Error GoTo Handler
        CommonDialog1.CancelError = True
        CommonDialog1.FileName = ""
        CommonDialog1.Filter = "all files (*.*)|*.*|bmp files (*.bmp)|*.bmp|jpeg files (*.jpg)|*.jpg|ico files (*.ico)|*.ico|gif files (*.gif)|*.gif"
        CommonDialog1.ShowOpen
        
        lblPicHolder.Caption = CommonDialog1.FileName
        imgPic.Picture = LoadPicture(lblPicHolder.Caption)
    
Handler:
    End If
End Sub

Private Sub lstParty_Click()
    lblParty = lstParty
End Sub

Private Sub lstParty_DblClick()
    Call cmdEnter_Click
End Sub

Private Sub lstPos_Click()
    
    ChangeCatcher
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As Candidates
    
    NameToSearch = lstPos
    If NameToSearch = "" Then
        Exit Sub
    End If
    
    NameToSearch = UCase(NameToSearch)
    Found = False
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, TmpPerson
        If NameToSearch = UCase(Trim(TmpPerson.CPos)) Then
            If UCase(lblPartyDisplay.Caption) = UCase(Trim(TmpPerson.CParty)) Then
                Found = True
                Exit For
            End If
        End If
    Next
    If Found = True Then
        gCurrentRecord = RecNum
        cmdNew.Enabled = False
        ShowCandi
    Else
        MsgBox "Candidate for " + NameToSearch + " in party " + lblPartyDisplay.Caption + " has not been entered yet.", vbOKOnly + vbExclamation, "Search"
        cmdNew.Enabled = True
        Call cmdNew_Click
    End If
End Sub

Private Sub Timer1_Timer()
      lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
End Sub
Private Sub txtAPF_GotFocus()
    txtAPF.SelStart = 0
    txtAPF.SelLength = Len(txtAPF.Text)

End Sub

Private Sub txtAPL_GotFocus()
    txtAPL.SelStart = 0
    txtAPL.SelLength = Len(txtAPL.Text)

End Sub

Private Sub txtAPM_GotFocus()
    txtAPM.SelStart = 0
    txtAPM.SelLength = Len(txtAPM.Text)
End Sub

Private Sub txtAPROF_GotFocus()
    txtAPROF.SelStart = 0
    txtAPROF.SelLength = Len(txtAPROF.Text)

End Sub

Private Sub txtAPROL_GotFocus()
    txtAPROL.SelStart = 0
    txtAPROL.SelLength = Len(txtAPROL.Text)
End Sub

Private Sub txtAPROM_GotFocus()
    txtAPROM.SelStart = 0
    txtAPROM.SelLength = Len(txtAPROM.Text)
End Sub

Private Sub txtASF_GotFocus()
    txtASF.SelStart = 0
    txtASF.SelLength = Len(txtASF.Text)
End Sub

Private Sub txtASL_GotFocus()
    txtASL.SelStart = 0
    txtASL.SelLength = Len(txtASL.Text)
End Sub

Private Sub txtASM_GotFocus()
    txtASM.SelStart = 0
    txtASM.SelLength = Len(txtASM.Text)
End Sub

Private Sub txtATF_GotFocus()
    txtATF.SelStart = 0
    txtATF.SelLength = Len(txtATF.Text)

End Sub

Private Sub txtATL_GotFocus()
    txtATL.SelStart = 0
    txtATL.SelLength = Len(txtATL.Text)
End Sub

Private Sub txtATM_GotFocus()
    txtATM.SelStart = 0
    txtATM.SelLength = Len(txtATM.Text)
End Sub
Private Sub txtAVPF_GotFocus()
    txtAVPF.SelStart = 0
    txtAVPF.SelLength = Len(txtAVPF.Text)
End Sub
Private Sub txtAVPL_GotFocus()
    txtAVPL.SelStart = 0
    txtAVPL.SelLength = Len(txtAVPL.Text)
End Sub
Private Sub txtAVPM_GotFocus()
    txtAVPM.SelStart = 0
    txtAVPM.SelLength = Len(txtAVPM.Text)
End Sub
Private Sub txtPartyA_GotFocus()
    txtPartyA.SelStart = 0
    txtPartyA.SelLength = Len(txtPartyA.Text)

End Sub

Private Sub ListOfParty()
Dim gPersonX As Parties
Dim gFileNumX As Integer
Dim gRecordLenX As Long
Dim gCurrentRecordX As Long
Dim gLastRecordX As Long
Dim NewRecX As Boolean
Dim NewFileX As Boolean
Dim RecNumX As Integer

    gRecordLenX = Len(gPersonX)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Party.dat" For Random As gFileNumX Len = gRecordLenX

    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Party.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    lstParty.Clear
    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
        lstParty.AddItem UCase(Trim(gPersonX.Party))
        If RecNumX = 1 Then
            PartyHold = UCase(Trim(gPersonX.Party))
        End If
    Next
    
    Close #gFileNumX
End Sub


Private Sub ListOfPositions()
    Dim gPersonX As Positions
    Dim gFileNumX As Integer
    Dim gRecordLenX As Long
    Dim gCurrentRecordX As Long
    Dim gLastRecordX As Long
    Dim NewRecX As Boolean
    Dim NewFileX As Boolean
    Dim RecNumX As Integer
    
    gRecordLenX = Len(gPersonX)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Positions.dat" For Random As gFileNumX Len = gRecordLenX

    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Positions.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    lstPos.Clear
    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
        lstPos.AddItem UCase(Trim(gPersonX.Position))
        If RecNumX = 1 Then
            PosHold = UCase(Trim(gPersonX.Position))
        End If
    Next
    
    Close #gFileNumX
End Sub

Private Sub ListOfCourse()
    Dim gPersonX As Courses
    Dim gFileNumX As Integer
    Dim gRecordLenX As Long
    Dim gCurrentRecordX As Long
    Dim gLastRecordX As Long
    Dim NewRecX As Boolean
    Dim NewFileX As Boolean
    Dim RecNumX As Integer
    Dim Holder As String
    
    gRecordLenX = Len(gPersonX)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Course.dat" For Random As gFileNumX Len = gRecordLenX

    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Course.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    cboCCourse.Clear
    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
'to determine if there are already courses to choose from (if none, exit)
        txtHolder.Text = UCase(Trim(gPersonX.Course))
        Holder = txtHolder.Text
        If Holder = "" Then
            MsgBox "You cannot enter candidates without entering courses first.  Exit then go to Course Entry / Voter Statistics first.", vbOKOnly + vbExclamation, "Error"
            Close #gFileNumX
            Call cmdCandiExit_Click
            Exit For
            Exit Sub
        End If
        cboCCourse.AddItem UCase(Trim(gPersonX.Course))
    Next
    
    Close #gFileNumX
End Sub

Private Sub ListOfCandi()
    lstCandi.Clear
    For RecNumX = 1 To gLastRecord
        Get #gFileNum, RecNumX, gPerson
        lstCandi.AddItem UCase(Trim(gPerson.CFullName))
    Next
    
End Sub



Private Sub txtCFirstName_GotFocus()
    txtCFirstName.SelStart = 0
    txtCFirstName.SelLength = Len(txtCFirstName.Text)
End Sub
Private Sub txtCLastName_GotFocus()
    txtCLastName.SelStart = 0
    txtCLastName.SelLength = Len(txtCLastName.Text)
End Sub

Private Sub txtCMidName_GotFocus()
    txtCMidName.SelStart = 0
    txtCMidName.SelLength = Len(txtCMidName.Text)
End Sub

Private Sub txtCStudNo_GotFocus()
    txtCStudNo.SelStart = 0
    txtCStudNo.SelLength = Len(txtCStudNo.Text)
End Sub


