VERSION 5.00
Begin VB.Form frmCanvassVote 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vote Results"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      Height          =   7095
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Height          =   3375
         Left            =   8160
         TabIndex        =   7
         Top             =   3480
         Width           =   1695
         Begin VB.CommandButton cmdIndividual 
            BackColor       =   &H00808080&
            Caption         =   "View Each Candidate Votes"
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
         Begin VB.CommandButton cmdExit 
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
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdWinners 
            BackColor       =   &H00808080&
            Caption         =   "View Winning Candidates"
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
            TabIndex        =   0
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   10095
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VOTING STATUS"
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
         TabIndex        =   5
         Top             =   240
         Width           =   2925
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6600
      Top             =   7080
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
      Left            =   7200
      TabIndex        =   3
      Top             =   7200
      Width           =   390
   End
End
Attribute VB_Name = "frmCanvassVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCanvassExit_Click()
    frmAdministrator.Show
    Unload Me
    
End Sub




Private Sub Form_Load()
    'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    ShowCandi
    ShowVotes
    cmdCanvassExit.Default = True
    
End Sub
Private Sub ShowCandi()
'to show the voting status count
    Dim gPerson As Candidates
    Dim gFileNumX As Integer
    Dim gRecordLen As Long
    Dim gCurrentRecord As Long
    Dim gLastRecord As Long
    Dim MySys As New FileSystemObject
    Dim Changes As Boolean
    
    gRecordLen = Len(gPerson)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Candi.dat" For Random As gFileNumX Len = gRecordLen

    gCurrentRecord = 1
    gLastRecord = FileLen(App.Path + "\VOTING SYSTEM\Candi.dat") / gRecordLen
    
    If gLastRecord = 0 Then
        gLastRecord = 1
    End If
    
    Changes = False
    
    
    Get #gFileNumX, gCurrentRecord, gPerson
    
    'party A
    fmePartyAVotes.Caption = Trim(gPerson.APartyName)
    
    lblAPres.Caption = UCase(Trim(gPerson.APresFirstName) + " " + Left(Trim(gPerson.APresMidName), 1) + ". " + Trim(gPerson.APresLastName))
    lblAVPres.Caption = UCase(Trim(gPerson.AVPresFirstName) + " " + Left(Trim(gPerson.AVPresMidName), 1) + ". " + Trim(gPerson.AVPresLastName))
    lblASec.Caption = UCase(Trim(gPerson.ASecFirstName) + " " + Left(Trim(gPerson.ASecMidName), 1) + ". " + Trim(gPerson.ASecLastName))
    lblATrea.Caption = UCase(Trim(gPerson.ATreaFirstName) + " " + Left(Trim(gPerson.ATreaMidName), 1) + ". " + Trim(gPerson.ATreaLastName))
    lblAPro.Caption = UCase(Trim(gPerson.APROFirstName) + " " + Left(Trim(gPerson.APROMidName), 1) + ". " + Trim(gPerson.APROLastName))
    
    'party B
    fmePartyBVotes.Caption = Trim(gPerson.BPartyName)
    
    lblBPres.Caption = UCase(Trim(gPerson.BPresFirstName) + " " + Left(Trim(gPerson.BPresMidName), 1) + ". " + Trim(gPerson.BPresLastName))
    lblBVPres.Caption = UCase(Trim(gPerson.BVPresFirstName) + " " + Left(Trim(gPerson.BVPresMidName), 1) + ". " + Trim(gPerson.BVPresLastName))
    lblBSec.Caption = UCase(Trim(gPerson.BSecFirstName) + " " + Left(Trim(gPerson.BSecMidName), 1) + ". " + Trim(gPerson.BSecLastName))
    lblBTrea.Caption = UCase(Trim(gPerson.BTreaFirstName) + " " + Left(Trim(gPerson.BTreaMidName), 1) + ". " + Trim(gPerson.BTreaLastName))
    lblBPro.Caption = UCase(Trim(gPerson.BPROFirstName) + " " + Left(Trim(gPerson.BPROMidName), 1) + ". " + Trim(gPerson.BPROLastName))

    Close #gFileNumX
End Sub

Public Sub ShowVotes()
'to show the number of votes
    Dim gPersonC As Votes
    Dim gFileNumC As Integer
    Dim gRecordLenC As Long
    Dim gCurrentRecordC As Long
    Dim gLastRecordC As Long
    Dim MySysC As New FileSystemObject
    
    gRecordLenC = Len(gPersonC)
    gFileNumC = FreeFile
    Open App.Path + "\VOTING SYSTEM\Votes.dat" For Random As gFileNumC Len = gRecordLenC
    gCurrentRecordC = 1
    gLastRecordC = FileLen(App.Path + "\VOTING SYSTEM\Votes.dat") / gRecordLenC
    If gLastRecordC = 0 Then
        gLastRecordC = 1
    End If
    
    Get #gFileNumC, gCurrentRecordC, gPersonC
    
    lblAPresVotes.Caption = gPersonC.PresA
    lblBPresVotes.Caption = gPersonC.PresB
 
    lblAVPresVotes.Caption = gPersonC.VPresA
    lblBVPresVotes.Caption = gPersonC.VPresB
 
    lblASecVotes.Caption = gPersonC.SecA
    lblBSecVotes.Caption = gPersonC.SecB
     
    lblATreaVotes.Caption = gPersonC.TreaA
    lblBTreaVotes.Caption = gPersonC.TreaB
 
    lblAProVotes.Caption = gPersonC.PROA
    lblBProVotes.Caption = gPersonC.PROB
    
    lblATotalVotes.Caption = Val(lblAPresVotes.Caption) + Val(lblAVPresVotes.Caption) + Val(lblASecVotes.Caption) + Val(lblATreaVotes.Caption) + Val(lblAProVotes.Caption)
    lblBTotalVotes.Caption = Val(lblBPresVotes.Caption) + Val(lblBVPresVotes.Caption) + Val(lblBSecVotes.Caption) + Val(lblBTreaVotes.Caption) + Val(lblBProVotes.Caption)
    
    Close #gFileNumC
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #gFileNumX
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
End Sub
