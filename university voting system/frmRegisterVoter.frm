VERSION 5.00
Begin VB.Form frmVoterRegistration 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VOTER REGISTRATION"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmeOpening 
      BackColor       =   &H80000007&
      Height          =   4935
      Left            =   120
      TabIndex        =   29
      Top             =   360
      Width           =   10455
      Begin VB.Frame Frame3 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   9975
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "REGISTRATION"
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
            TabIndex        =   33
            Top             =   240
            Width           =   2715
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Height          =   1335
         Left            =   8520
         TabIndex        =   30
         Top             =   3360
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
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Image imgLogo 
         Height          =   2895
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   3255
      End
   End
   Begin VB.Frame fmeStart 
      BackColor       =   &H80000007&
      Height          =   5295
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtHolder 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   4800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00000000&
         Caption         =   "Registered Voters"
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
         Height          =   4215
         Left            =   5760
         TabIndex        =   27
         Top             =   960
         Width           =   2775
         Begin VB.ListBox lstVoters 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3840
            ItemData        =   "frmRegisterVoter.frx":0000
            Left            =   120
            List            =   "frmRegisterVoter.frx":0007
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fmeRegister 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   4215
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   5535
         Begin VB.Frame Frame10 
            BackColor       =   &H80000007&
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   5295
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
               TabIndex        =   20
               Top             =   240
               Width           =   60
            End
         End
         Begin VB.TextBox txtVStudNum 
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox cboVCourse 
            Height          =   315
            ItemData        =   "frmRegisterVoter.frx":0016
            Left            =   1200
            List            =   "frmRegisterVoter.frx":0032
            TabIndex        =   5
            Top             =   3360
            Width           =   4095
         End
         Begin VB.TextBox txtVFirstName 
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   2400
            Width           =   2895
         End
         Begin VB.TextBox txtVMidName 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   2760
            Width           =   2895
         End
         Begin VB.TextBox txtVLastName 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label lblIdentityValidator 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   3600
            Visible         =   0   'False
            Width           =   4935
         End
         Begin VB.Label Label6 
            BackColor       =   &H00000000&
            Caption         =   "Student No:"
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
            Height          =   225
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H00000000&
            Caption         =   "Course:"
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
            Height          =   225
            Left            =   120
            TabIndex        =   24
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Caption         =   "Firstname:"
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
            Height          =   300
            Left            =   120
            TabIndex        =   23
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            Caption         =   "Middlename:"
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
            Height          =   300
            Left            =   120
            TabIndex        =   22
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Caption         =   "Lastname:"
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
            Height          =   225
            Left            =   120
            TabIndex        =   21
            Top             =   2040
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   8415
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "REGISTRATION"
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
            TabIndex        =   17
            Top             =   240
            Width           =   2715
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000012&
         Height          =   5055
         Left            =   8640
         TabIndex        =   15
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton cmdDelAll 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Delete &All"
            Enabled         =   0   'False
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
            TabIndex        =   10
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdExit 
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
            Height          =   855
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmRegisterVoter.frx":0062
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdDel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Delete"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmRegisterVoter.frx":04A4
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2760
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
            Height          =   855
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmRegisterVoter.frx":05A6
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Register"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmRegisterVoter.frx":0C10
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Edit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmRegisterVoter.frx":127A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1920
            Width           =   1455
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   6240
   End
   Begin VB.Label lblTimeTrack 
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
      Left            =   120
      TabIndex        =   34
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label lblRecNum 
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
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
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
      Left            =   7680
      TabIndex        =   12
      Top             =   5640
      Width           =   390
   End
End
Attribute VB_Name = "frmVoterRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gPerson As VoterInfo
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim gCurrentRecord As Long
Dim gLastRecord As Long
Dim MySys As New FileSystemObject
Dim NewRec As Boolean
Dim NewFile As Boolean
Dim Saw As Boolean
Private Sub SaveVoter()
'to save the registration of a voter
    gPerson.VLastName = txtVLastName.Text
    gPerson.VFirstName = txtVFirstName.Text
    gPerson.VMidName = txtVMidName.Text
    gPerson.VCourse = cboVCourse.Text
    gPerson.VStudNum = txtVStudNum.Text
    gPerson.VValidator = lblIdentityValidator.Caption
    gPerson.VFullName = UCase(txtVLastName.Text + ", " + txtVFirstName.Text + " " + Left(txtVMidName.Text, 1) + ".")
    gPerson.VTimeTrack = lblTime.Caption
  
    Put #gFileNum, gCurrentRecord, gPerson
    NewRec = False
    fmeRegister.Enabled = False
    Call ListOfVoters
End Sub
Private Sub RegError()
    MsgBox "Fill up the registration COMPLETELY!", vbOKOnly + vbExclamation
End Sub

Private Sub cmdDel_Click()
   If lblName = "" Then
        MsgBox "There are no records to delete.", vbOKOnly + vbExclamation, "Delete"
        Exit Sub
    End If
    
    Dim DirResult
    Dim TmpFileNum
    Dim TmpPerson As VoterInfo
    Dim RecNum As Long
    Dim TmpRecNum As Long
    Dim Msg As String
    
    Msg = "Delete " + UCase(txtVFirstName.Text + " " + txtVLastName.Text) + "'s record?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbNo Then
        lstVoters = lblName.Caption
        lstVoters.SetFocus
        Exit Sub
    End If
    
    
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\MyVoters.tmp") = True Then
        Kill App.Path + "\VOTING SYSTEM\MyVoters.tmp"
    End If
    TmpFileNum = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\MyVoters.tmp" For Random As TmpFileNum Len = gRecordLen
        
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
    
    
    MySys.DeleteFile App.Path + "\VOTING SYSTEM\Voters.dat"
    
    Close #TmpFileNum
    Name App.Path + "\VOTING SYSTEM\MyVoters.tmp" As App.Path + "\VOTING SYSTEM\Voters.dat"
    
    gFileNum = FreeFile
    Open App.Path + "\VOTING SYSTEM\Voters.dat" For Random As gFileNum Len = gRecordLen
    
    gLastRecord = gLastRecord - 1
    If gLastRecord = 0 Then gLastRecord = 1
    
    If gCurrentRecord > gLastRecord Then
        gCurrentRecord = gLastRecord
    End If
    
    MsgBox "Voter deleted.", vb0konly + vbInformation, "Delete"
    Call ListOfVoters
    
    ShowVoter
    lstVoters = lblName.Caption
    lstVoters.SetFocus
    If lblName.Caption = "" Then
       NewFile = True
       lblRecNum.Visible = False
    Else
       NewRec = False
    End If
    
End Sub

Private Sub cmdDelAll_Click()
    If lstVoters = "" Then
        MsgBox "There are no records to delete.", vbOKOnly + vbExclamation, "Delete"
        Exit Sub
    End If
    
    Msg = "Are you sure you want ot delete all records?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbYes Then
        If MySys.FileExists(App.Path + "\VOTING SYSTEM\Voters.dat") = True Then
            Close #gFileNum
            Kill App.Path + "\VOTING SYSTEM\Voters.dat"
            
            gFileNum = FreeFile
            Open App.Path + "\VOTING SYSTEM\Voters.dat" For Random As gFileNum Len = gRecordLen
    
            gLastRecord = 1
            gCurrentRecord = gLastRecord
            

            Call ListOfVoters
              
            ShowVoter
            MsgBox "All voter records cleared!", vb0konly + vbInformation, "Delete"
            lstVoters = lblName.Caption
            lstVoters.SetFocus
            If lblName.Caption = "" Then
                NewFile = True
                lblRecNum.Visible = False
            Else
                NewRec = False
            End If
        End If
    End If
    
    lstVoters = lblName.Caption
    lstVoters.SetFocus
End Sub

Private Sub cmdEdit_Click()
    fmeRegister.Enabled = True
    cmdSave.Enabled = True
    cmdSave.Default = True
    txtVStudNum.SetFocus
End Sub

Private Sub cmdEnter_Click()
    
    fmeOpening.Visible = False
    fmeStart.Visible = True
    cmdDel.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdDelAll.Enabled = False
    fmeRegister.Enabled = False
    cmdNew.Default = True
    StartReg
End Sub

Private Sub cmdEnter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        ToPass = 3
        frmPassAdmin.Show
        Me.Enabled = False
    End If
End Sub

Private Sub cmdExit_Click()
    ChangeCatcher
    Close #gFileNum
    fmeStart.Visible = False
    fmeOpening.Visible = True
    cmdEnter.Default = True
    cmdEnter.SetFocus
    lblRecNum.Visible = False
    lblTimeTrack.Visible = False
End Sub

Private Sub cmdExit_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        pass = InputBox("Enter COMMAND LINE (passwordS) to allow deletion and editng of voter records.", "Password")
        If pass = "show me the money" Then
            MsgBox "Access Granted!", vbOKOnly + vbInformation, "Password"
            cmdDel.Enabled = True
            cmdDelAll.Enabled = True
            cmdEdit.Enabled = True
        ElseIf pass = "" Then
        Else
            MsgBox "Access Denied!", vbOKOnly + vbExclamation, "Password"
            cmdNew.SetFocus
        End If
    End If
    
End Sub

Private Sub cmdNew_Click()
    Call ChangeCatcher
    fmeRegister.Enabled = True
    cmdSave.Enabled = True
    cmdSave.Default = True
    NewRec = True
    lblRecNum.Visible = False
    lblName.Caption = ""
    lblTimeTrack.Visible = False
    
    gPerson.VCourse = ""
    gPerson.VFirstName = ""
    gPerson.VFullName = ""
    gPerson.VLastName = ""
    gPerson.VMidName = ""
    gPerson.VStudNum = ""
    gPerson.VTimeTrack = ""
    gPerson.VVoteTrack = 0
    gPerson.VVote = ""
    
    
    txtVLastName.Text = ""
    txtVFirstName.Text = ""
    txtVMidName.Text = ""
    cboVCourse.Text = ""
    txtVStudNum.Text = ""
    lblIdentityValidator.Caption = ""
    
    cboVCourse.Tag = ""
    txtVLastName.Tag = ""
    txtVMidName.Tag = ""
    txtVFirstName.Tag = ""
    txtVStudNum.Tag = ""
    
    txtVStudNum.SetFocus
    
End Sub

Private Sub cmdSave_Click()
'to avoid saving incomplete data
    CourseValidator
    If Saw = False Then
        MsgBox "You entered an invalid course.  Choose only from the choices!", vbOKOnly + vbExclamation, "Save"
        cboVCourse.Text = ""
        cboVCourse.SetFocus
        Saw = False
        Exit Sub
    End If
    If txtVStudNum.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtVStudNum.SetFocus
        Exit Sub
    End If
    If cboVCourse.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        cboVCourse.SetFocus
        Exit Sub
    End If
    If txtVLastName.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtVLastName.SetFocus
        Exit Sub
    End If
    If txtVFirstName.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtVFirstName.SetFocus
        Exit Sub
    End If
    If txtVMidName.Text = "" Then
        MsgBox "Fill up entry completely before saving", vbOKOnly + vbExclamation, "Error Save"
        txtVMidName.SetFocus
        Exit Sub
    End If

'to determine if registration is valid
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As VoterInfo
    Dim StudNumToSearch As String
    
    lblIdentityValidator.Caption = UCase(txtVLastName.Text + ", " + txtVFirstName.Text + ", " + txtVMidName.Text + " " + txtVStudNum.Text)
    NameToSearch = lblIdentityValidator.Caption
    StudNumToSearch = txtVStudNum.Text
    Found = False
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, TmpPerson
        If NewRec = True Then
            If NameToSearch = Trim(TmpPerson.VValidator) Or StudNumToSearch = Trim(TmpPerson.VStudNum) Then
                gCurrentRecord = RecNum
                Found = True
                Exit For
            End If
        End If
    Next
    If Found = True Then
        MsgBox "Error in registration. You may be registering more than once or you entered a duplicate student number.", vbOKOnly + vbExclamation, "Error Registration"
        txtVStudNum.SetFocus
        Exit Sub
    Else
        If NewRec = True And NewFile = False Then
            gLastRecord = gLastRecord + 1
            Put #gFileNum, gLastRecord, gPerson
            gCurrentRecord = gLastRecord
            SaveVoter
           
            MsgBox "Voter now registered.", vbOKOnly + vbInformation, "Save"
            cmdSave.Enabled = False
            cmdNew.Default = True
            ShowVoter
        Else
            SaveVoter
            MsgBox "Voter now registered.", vbOKOnly + vbInformation, "Save"
            cmdSave.Enabled = False
            cmdNew.Default = True
            ShowVoter
            NewFile = False
        End If
    End If
    On Error GoTo 100
     
    lstVoters = lblName.Caption
    lstVoters.SetFocus
    Exit Sub
100
    lstVoters = lblName.Caption
    lstVoters.SetFocus
End Sub

Private Sub Form_Load()
'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    imgLogo.Picture = LoadPicture(MyPicture)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #gFileNum
End Sub

Private Sub lstVoters_Click()
    ChangeCatcher
    fmeRegister.Enabled = False
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As VoterInfo
    
    NameToSearch = lstVoters
    If NameToSearch = "" Then
        Exit Sub
    End If
    
    NameToSearch = UCase(NameToSearch)
    Found = False
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, TmpPerson
        If NameToSearch = UCase(Trim(TmpPerson.VFullName)) Then
            Found = True
            Exit For
        End If
    Next
    If Found = True Then
        gCurrentRecord = RecNum
        ShowVoter
    Else
        MsgBox "Voter " + NameToSearch + " not found.", vbOKOnly + vbExclamation, "Search"
    End If
    
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
End Sub
Private Sub txtFirstName_GotFocus()
    txtVFirstName.SelStart = 0
    txtVFirstName.SelLength = Len(txtVFirstName)
End Sub
Private Sub txtLastName_GotFocus()
    txtVLastName.SelStart = 0
    txtVLastName.SelLength = Len(txtVLastName)
End Sub
Private Sub txtMidName_GotFocus()
    txtVMidName.SelStart = 0
    txtVMidName.SelLength = Len(txtVMidName)
End Sub
Private Sub txtStudNum_GotFocus()
    txtVStudNum.SelStart = 0
    txtVStudNum.SelLength = Len(txtVStudNum)
End Sub

Private Sub ChangeCatcher()
    
    If txtVLastName.Tag <> txtVLastName.Text Then
        GoTo SaveNow
    End If
    If txtVFirstName.Tag <> txtVFirstName.Text Then
        GoTo SaveNow
    End If
    If txtVMidName.Tag <> txtVMidName.Text Then
        GoTo SaveNow
    End If
    If txtVStudNum.Tag <> txtVStudNum.Text Then
        GoTo SaveNow
    End If
    Exit Sub
SaveNow:
        If lblName.Caption = "" Then
            Msg = "Save changes to this new record?"
        Else
            Msg = "Save changes to the record of " + UCase(lblName.Caption) + "?"
        End If
        Response = MsgBox(Msg, vbYesNo + vbQuestion, "Save" & Save)
        If Response = vbYes Then
            Call cmdSave_Click
        Else
            Exit Sub
        End If
End Sub



Private Sub ListOfVoters()
    lstVoters.Clear
    For RecNumX = 1 To gLastRecord
        Get #gFileNum, RecNumX, gPerson
        lstVoters.AddItem UCase(Trim(gPerson.VFullName))
    Next
End Sub

Private Sub ShowVoter()
'to display the current voters registered
    Get #gFileNum, gCurrentRecord, gPerson
    
    txtVLastName.Text = Trim(gPerson.VLastName)
    txtVFirstName.Text = Trim(gPerson.VFirstName)
    txtVMidName.Text = Trim(gPerson.VMidName)
    txtVStudNum.Text = Trim(gPerson.VStudNum)
    cboVCourse.Text = Trim(gPerson.VCourse)
    
    lblName.Caption = Trim(gPerson.VFullName)
    lblTimeTrack.Caption = "Registration saved last: " + Trim(gPerson.VTimeTrack) + "."
    
    txtVLastName.Tag = txtVLastName.Text
    txtVFirstName.Tag = txtVFirstName.Text
    txtVMidName.Tag = txtVMidName.Text
    txtVStudNum.Tag = txtVStudNum.Text
    cboVCourse.Tag = cboVCourse.Text
    
    lblRecNum.Visible = True
    lblTimeTrack.Visible = True
    lblRecNum.Caption = "Voter #" + _
                    Str(gCurrentRecord) + " of " + _
                    Str(gLastRecord) + " saved voter(s)."
    NewRec = False
    fmeRegister.Enabled = False
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
    
    gRecordLenX = Len(gPersonX)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Course.dat" For Random As gFileNumX Len = gRecordLenX

    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Course.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    cboVCourse.Clear
    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
'to determine if there are already courses to choose from (if none, exit)
        txtHolder.Text = UCase(Trim(gPersonX.Course))
        Holder = txtHolder.Text
        If Holder = "" Then
            MsgBox "You cannot register without entering courses first.  Ask for Administrator assistance.", vbOKOnly + vbExclamation, "Error"
            Close #gFileNumX
            Call cmdExit_Click
            Exit For
            Exit Sub
        End If
        cboVCourse.AddItem UCase(Trim(gPersonX.Course))
    Next
    
    Close #gFileNumX
End Sub
Private Sub txtvFirstName_GotFocus()
    txtVFirstName.SelStart = 0
    txtVFirstName.SelLength = Len(txtVFirstName.Text)
End Sub
Private Sub txtvLastName_GotFocus()
    txtVLastName.SelStart = 0
    txtVLastName.SelLength = Len(txtVLastName.Text)
End Sub

Private Sub txtvMidName_GotFocus()
    txtVMidName.SelStart = 0
    txtVMidName.SelLength = Len(txtVMidName.Text)
End Sub

Private Sub txtvStudNum_GotFocus()
    txtVStudNum.SelStart = 0
    txtVStudNum.SelLength = Len(txtVStudNum.Text)
End Sub


Private Sub StartReg()
    'to determine if there are no currently saved voters
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\Voters.dat") = True Then
        NewFile = False
    Else
        NewFile = True
        MsgBox "No currently registered voters!", vbOKOnly + vbInformation, "Candidates"
    End If
   

'to open the VOTERS DATABASE
    gRecordLen = Len(gPerson)

    gFileNum = FreeFile

    Open App.Path + "\VOTING SYSTEM\Voters.dat" For Random As gFileNum Len = gRecordLen

    gCurrentRecord = 1
    gLastRecord = FileLen(App.Path + "\VOTING SYSTEM\Voters.dat") / gRecordLen
  
    If gLastRecord = 0 Then
        gLastRecord = 1
        NewFile = True
    End If
    Call ListOfVoters
    ShowVoter
    lstVoters = lblName.Caption
    Call ListOfCourse
End Sub

Private Sub CourseValidator()
'to determine if the course entered is on the saved course list
    Dim gPersonX As Courses
    Dim gFileNumX As Integer
    Dim gRecordLenX As Long
    Dim gCurrentRecordX As Long
    Dim gLastRecordX As Long
    Dim NewRecX As Boolean
    Dim NewFileX As Boolean
    Dim RecNumX As Integer
    
    gRecordLenX = Len(gPersonX)

    gFileNumX = FreeFile

    Open App.Path + "\VOTING SYSTEM\Course.dat" For Random As gFileNumX Len = gRecordLenX

    gCurrentRecordX = 1
    gLastRecordX = FileLen(App.Path + "\VOTING SYSTEM\Course.dat") / gRecordLenX
    
    If gLastRecordX = 0 Then
        gLastRecordX = 1
    End If
    
    Saw = False
    
    For RecNumX = 1 To gLastRecordX
        Get #gFileNumX, RecNumX, gPersonX
        If UCase(Trim(gPersonX.Course)) = cboVCourse.Text Then
            Saw = True
            Exit For
        End If
    Next
    
    Close #gFileNumX
End Sub
