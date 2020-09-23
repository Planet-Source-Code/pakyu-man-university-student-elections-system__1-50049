VERSION 5.00
Begin VB.Form frmVoteResults 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CANDIDATES"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmeOpening 
      BackColor       =   &H80000007&
      Height          =   4575
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   11655
      Begin VB.Frame Frame8 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   9615
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "VOTE RESULTS"
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
            TabIndex        =   41
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00000000&
         Height          =   4335
         Left            =   9840
         TabIndex        =   36
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton cmdAllResults 
            BackColor       =   &H00808080&
            Caption         =   "View Summary of &All Candidates' Standing"
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
            TabIndex        =   38
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdWinners 
            BackColor       =   &H00808080&
            Caption         =   "View &Winning Candidates Summary"
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
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdExitResults 
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
            TabIndex        =   42
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CommandButton cmdIndi 
            BackColor       =   &H00808080&
            Caption         =   "View &Individual Votes Garnered"
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
            TabIndex        =   40
            Top             =   2160
            Width           =   1455
         End
      End
      Begin VB.Image imgLogo 
         Height          =   2655
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   3015
      End
   End
   Begin VB.Frame fmeSummary 
      BackColor       =   &H80000007&
      Height          =   5415
      Left            =   2400
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox txtResults 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   43
         Top             =   960
         Width           =   5055
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00000000&
         Height          =   2415
         Left            =   5520
         TabIndex        =   32
         Top             =   2760
         Width           =   1695
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00808080&
            Caption         =   "&Print"
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
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdQuit 
            BackColor       =   &H00808080&
            Caption         =   "&Quit"
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
            TabIndex        =   33
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   240
         TabIndex        =   30
         Top             =   120
         Width           =   6975
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "VOTE RESULTS SUMMARY"
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
            TabIndex        =   31
            Top             =   240
            Width           =   4725
         End
      End
   End
   Begin VB.Frame fmeStart 
      BackColor       =   &H80000007&
      Height          =   4575
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame7 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   11415
         Begin VB.Label lblParty 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   60
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Height          =   2415
         Left            =   9840
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
         Begin VB.CommandButton cmdEnter 
            BackColor       =   &H00808080&
            Caption         =   "&Enter"
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
         Begin VB.CommandButton cmdQuitIndi 
            BackColor       =   &H00808080&
            Caption         =   "&Quit"
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
         TabIndex        =   19
         Top             =   1680
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
            Left            =   120
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
         TabIndex        =   23
         Top             =   1320
         Width           =   4215
      End
   End
   Begin VB.Frame fmeEntry 
      BackColor       =   &H80000007&
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame13 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   11415
         Begin VB.Label lblCourse 
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
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   45
         End
      End
      Begin VB.TextBox txtHolder 
         Height          =   285
         Left            =   0
         TabIndex        =   24
         Top             =   5040
         Visible         =   0   'False
         Width           =   3015
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
         Height          =   3135
         Left            =   6720
         TabIndex        =   17
         Top             =   2040
         Width           =   3015
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
            Height          =   2760
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   6495
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
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   45
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   11415
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "INDIVIDUAL VOTE CANVASS"
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
            TabIndex        =   14
            Top             =   120
            Width           =   4980
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000012&
         Height          =   1335
         Left            =   9840
         TabIndex        =   12
         Top             =   3840
         Width           =   1695
         Begin VB.CommandButton cmdResultsIndiBack 
            BackColor       =   &H00808080&
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
            Height          =   975
            Left            =   120
            MaskColor       =   &H8000000A&
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   11415
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
            TabIndex        =   11
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   6495
         Begin VB.Label lblName 
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
            TabIndex        =   9
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "NUMBER OF VOTES GARNERED:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label lblVotes 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   195
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
      TabIndex        =   6
      Top             =   5520
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
      TabIndex        =   5
      Top             =   5520
      Width           =   390
   End
End
Attribute VB_Name = "frmVoteResults"
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
    Dim PartyHold As String

Private Sub cmdAllResults_Click()
    txtResults.Text = vbNewLine + "             E L E C T I O N S" + vbNewLine + "               R E S U L T S" + vbNewLine + vbNewLine
    Dim gPersonZ As Parties
    Dim gFileNumZ As Integer
    Dim gRecordLenZ As Long
    Dim gCurrentRecordZ As Long
    Dim gLastRecordZ As Long
    Dim RecNumZ As Integer
    
    gRecordLenZ = Len(gPersonZ)
    
    gFileNumZ = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\Party.dat" For Random As gFileNumZ Len = gRecordLenZ
    
    gCurrentRecordZ = 1
    gLastRecordZ = FileLen(App.Path + "\VOTING SYSTEM\Party.dat") / gRecordLenZ
    
    If gLastRecordZ = 0 Then
        gLastRecordZ = 1
    End If
    
    
    
    Dim gPersonX As Candidates
    Dim gFileNumX As Integer
    Dim gRecordLenX As Long
    Dim gCurrentRecordX As Long
    Dim gLastRecordX As Long
    Dim RecNumX As Integer
    
    Dim num As Integer
    Dim Highest As String
    
    
    gRecordLenX = Len(gPersonX)
    
    gFileNumX = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNumX Len = gRecordLenX
    
    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Candi.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    

    For RecNumZ = 1 To gLastRecordZ
        
        Get #gFileNumZ, RecNumZ, gPersonZ
        txtResults.Text = txtResults.Text + " " + UCase(Trim(gPersonZ.Party)) + vbNewLine + vbNewLine
        For RecNumX = 1 To gLastRecordX
            Get #gFileNumX, RecNumX, gPersonX
            If UCase(Trim(gPersonX.CParty)) = UCase(Trim(gPersonZ.Party)) Then
                txtResults.Text = txtResults.Text + "   FOR " + UCase(Trim(gPersonX.CPos)) + vbNewLine + "     " + UCase(Trim(gPersonX.CFullName)) + vbNewLine + "     Course:  " + UCase(Trim(gPersonX.CCourse)) + vbNewLine + "     Total Votes Garnered:  " + UCase(Trim(gPersonX.CVote)) + vbNewLine + vbNewLine
            End If
        Next
        
    Next
    fmeSummary.Visible = True
    fmeOpening.Visible = False
    cmdPrint.Default = True
    txtResults.Text = txtResults.Text + vbNewLine + "  As of " + lblTime.Caption
    Close #gFileNumX
    Close #gFileNumZ
End Sub

Private Sub cmdEnter_Click()
    fmeStart.Visible = False
    fmeEntry.Visible = True
    Call ListOfCandidatesByParty
    
    lblPartyDisplay.Caption = lblParty.Caption
    cmdResultsIndiBack.Default = True
    ShowCandi
    lstPos = lblName.Caption
On Error GoTo JustGo
    lstPos.SetFocus
JustGo:
End Sub


Private Sub cmdExitResults_Click()
    frmAdministrator.Show
    Unload Me
End Sub

Private Sub cmdIndi_Click()
    OpenIndividualCandidates
    fmeOpening.Visible = False
    fmeStart.Visible = True
    lstParty.SetFocus
    cmdEnter.Default = True
    
End Sub




Private Sub cmdPrint_Click()
    Printer.Font = "courier new"
    Printer.Print txtResults.Text
    Printer.EndDoc
End Sub

Private Sub cmdQuit_Click()
    txtResults.Text = ""
    fmeSummary.Visible = False
    fmeOpening.Visible = True
    cmdWinners.Default = True
    
    
End Sub

Private Sub cmdQuitIndi_Click()
    fmeStart.Visible = False
    fmeOpening.Visible = True
    cmdWinners.Default = True
    cmdWinners.SetFocus
    Close #gFileNum
End Sub

Private Sub cmdResultsIndiBack_Click()
    fmeEntry.Visible = False
    fmeStart.Visible = True
    cmdEnter.Default = True
    lstParty.SetFocus
End Sub

Private Sub cmdWinners_Click()
    txtResults.Text = vbNewLine + "             E L E C T I O N S" + vbNewLine + "       W I N N I N G    R E S U L T S" + vbNewLine + vbNewLine
    Dim gPersonZ As Positions
    Dim gFileNumZ As Integer
    Dim gRecordLenZ As Long
    Dim gCurrentRecordZ As Long
    Dim gLastRecordZ As Long
    Dim RecNumZ As Integer
    
    gRecordLenZ = Len(gPersonZ)
    
    gFileNumZ = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\Positions.dat" For Random As gFileNumZ Len = gRecordLenZ
    
    gCurrentRecordZ = 1
    gLastRecordZ = FileLen(App.Path + "\VOTING SYSTEM\Positions.dat") / gRecordLenZ
    
    If gLastRecordZ = 0 Then
        gLastRecordZ = 1
    End If
    
    
    
    Dim gPersonX As Candidates
    Dim gFileNumX As Integer
    Dim gRecordLenX As Long
    Dim gCurrentRecordX As Long
    Dim gLastRecordX As Long
    Dim RecNumX As Integer
    
    Dim num As Integer
    Dim Highest As String
    
    
    gRecordLenX = Len(gPersonX)
    
    gFileNumX = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNumX Len = gRecordLenX
    
    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Candi.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    
    
    For RecNumZ = 1 To gLastRecordZ
        num = -1
        Get #gFileNumZ, RecNumZ, gPersonZ
   
        For RecNumX = 1 To gLastRecordX
            Get #gFileNumX, RecNumX, gPersonX
            If UCase(Trim(gPersonX.CPos)) = UCase(Trim(gPersonZ.Position)) Then
                If Val(gPersonX.CVote) > num Then
                    num = gPersonX.CVote
                    jam = txtResults.Text + " FOR " + UCase(Trim(gPersonX.CPos)) + vbNewLine + vbNewLine + "     " + UCase(Trim(gPersonX.CFullName)) + vbNewLine + "     Course:  " + UCase(Trim(gPersonX.CCourse)) + vbNewLine + "     Party:  " + UCase(Trim(gPersonX.CParty)) + vbNewLine + "     Total Votes Garnered:  " + UCase(Trim(gPersonX.CVote)) + vbNewLine + vbNewLine
                ElseIf Val(gPersonX.CVote) = num Then
                    num = gPersonX.CVote
                    jam = jam + "     T I E S   W I T H " + vbNewLine + vbNewLine + "     " + UCase(Trim(gPersonX.CFullName)) + vbNewLine + "     Party:  " + UCase(Trim(gPersonX.CParty)) + vbNewLine + "     Total Votes Garnered:  " + UCase(Trim(gPersonX.CVote)) + vbNewLine + vbNewLine
                End If
            
            End If
        Next
        txtResults.Text = jam
    Next
    fmeSummary.Visible = True
    fmeOpening.Visible = False
    cmdPrint.Default = True
    txtResults.Text = txtResults.Text + vbNewLine + "  As of " + lblTime.Caption
    Close #gFileNumX
    Close #gFileNumZ
End Sub

Private Sub Form_Load()
'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    imgLogo.Picture = LoadPicture(MyPicture)
End Sub

Private Sub ShowCandi()
'to display the candidates
    Get #gFileNum, gCurrentRecord, gPerson
    
    lblName.Caption = Trim(gPerson.CFullName)
    lblPosition.Caption = "FOR " + Trim(gPerson.CPos)
    lblCourse.Caption = Trim(gPerson.CCourse)
    
    lblVotes.Caption = gPerson.CVote
    
    
    lblRecNum.Visible = True
    lblRecNum.Caption = "Candidate #" + _
                    Str(gCurrentRecord) + " of " + _
                    Str(gLastRecord) + " saved candidate(s)."
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmAdministrator.Show
    Close #gFileNum
End Sub
Private Sub lstParty_Click()
    lblParty = lstParty
End Sub
Private Sub lstParty_DblClick()
    Call cmdEnter_Click
End Sub
Private Sub lstPos_Click()
    
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
        If NameToSearch = UCase(Trim(TmpPerson.CFullName)) Then
                Found = True
                Exit For
        End If
    Next
    If Found = True Then
        gCurrentRecord = RecNum
        ShowCandi
    End If
End Sub

Private Sub Timer1_Timer()
      lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
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
Private Sub ListOfCandidatesByParty()
    Dim gPersonX As Candidates
    Dim gFileNumX As Integer
    Dim gRecordLenX As Long
    Dim gCurrentRecordX As Long
    Dim gLastRecordX As Long
    Dim NewRecX As Boolean
    Dim NewFileX As Boolean
    Dim RecNumX As Integer
    Dim Found As Boolean
    Found = False
    gRecordLenX = Len(gPersonX)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNumX Len = gRecordLenX

    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Candi.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    lstPos.Clear
    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
        If UCase(Trim(gPersonX.CParty)) = lstParty Then
            lstPos.AddItem UCase(Trim(gPersonX.CFullName))
            Found = True
            gCurrentRecord = RecNumX
            
        End If
    Next
    If Found = False Then
        display = "There are no candidates entered in party " + lblParty.Caption + "."
        MsgBox display, vbOKOnly + vbExclamation, "No entry"
        Call cmdResultsIndiBack_Click
    End If
    
    Close #gFileNumX
End Sub
Private Sub ListOfCandi()
    lstCandi.Clear
    For RecNumX = 1 To gLastRecord
        Get #gFileNum, RecNumX, gPerson
        lstCandi.AddItem UCase(Trim(gPerson.CFullName))
    Next
End Sub

Private Sub OpenIndividualCandidates()
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
End Sub
