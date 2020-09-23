VERSION 5.00
Begin VB.Form frmVoterStats 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voter Statistics"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmeStart 
      BackColor       =   &H00000000&
      Height          =   7215
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   10935
      Begin VB.Frame Frame11 
         BackColor       =   &H80000007&
         Height          =   735
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   8655
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "COURSE ENTRY"
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
            TabIndex        =   34
            Top             =   240
            Width           =   2820
         End
      End
      Begin VB.TextBox txtCourse 
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Top             =   2880
         Width           =   6375
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   8655
         Begin VB.Label lblCourseB 
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
            TabIndex        =   31
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Courses Saved"
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
         Height          =   2775
         Left            =   240
         TabIndex        =   29
         Top             =   3360
         Width           =   8655
         Begin VB.ListBox lstCourse 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   2340
            ItemData        =   "frmVoterStats.frx":0000
            Left            =   120
            List            =   "frmVoterStats.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   8415
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Height          =   5295
         Left            =   9000
         TabIndex        =   28
         Top             =   840
         Width           =   1695
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Delete A&ll"
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
            TabIndex        =   5
            Top             =   3120
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
            Picture         =   "frmVoterStats.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Save Course"
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
            Picture         =   "frmVoterStats.frx":0446
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdDel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Delete Course"
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
            Picture         =   "frmVoterStats.frx":0AB0
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Add Course"
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
            Picture         =   "frmVoterStats.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
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
            Height          =   735
            Left            =   120
            MaskColor       =   &H8000000A&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   4440
            Width           =   1455
         End
         Begin VB.CommandButton cmdEnterVoterStats 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enter Statist&ics"
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
            Height          =   735
            Left            =   120
            MaskColor       =   &H8000000A&
            Picture         =   "frmVoterStats.frx":121C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3600
            Width           =   1455
         End
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
         TabIndex        =   36
         Top             =   6240
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label label1 
         BackColor       =   &H80000007&
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2640
         Width           =   3015
      End
   End
   Begin VB.Frame fmeFinal 
      BackColor       =   &H80000007&
      Height          =   3735
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   5895
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
         Height          =   3015
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   24
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "VOTE RESULTS"
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
         TabIndex        =   26
         Top             =   240
         Width           =   1740
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
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H80000007&
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   10935
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VOTER STATISTICS"
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
         TabIndex        =   22
         Top             =   240
         Width           =   3450
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   6120
      Width           =   9135
      Begin VB.Label lblVoteTime 
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
         TabIndex        =   38
         Top             =   720
         Width           =   45
      End
      Begin VB.Label lblRecordNum 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "time"
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
         TabIndex        =   37
         Top             =   240
         Width           =   375
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
         TabIndex        =   20
         Top             =   480
         Width           =   45
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   5880
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   10935
      Begin VB.Label lblStudNum 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "567"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "Student number:"
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
         TabIndex        =   16
         Top             =   120
         Width           =   1785
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000007&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   10935
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1140
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   9360
      TabIndex        =   10
      Top             =   4800
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
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Voters"
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
      Left            =   6120
      TabIndex        =   8
      Top             =   2400
      Width           =   3135
      Begin VB.ListBox lstRecords 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   3420
         ItemData        =   "frmVoterStats.frx":1886
         Left            =   120
         List            =   "frmVoterStats.frx":1888
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
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
      Left            =   7800
      TabIndex        =   18
      Top             =   7320
      Width           =   390
   End
   Begin VB.Label lblCourse 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Label2"
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
      TabIndex        =   14
      Top             =   840
      Width           =   540
   End
End
Attribute VB_Name = "frmVoterStats"
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

Dim ChosenCourse As String

Dim gPersonB As Courses
Dim gFileNumB As Integer
Dim gRecordLenB As Long
Dim gCurrentRecordB As Long
Dim gLastRecordB As Long

Dim NewRec As Boolean
Dim NewFile As Boolean


Private Sub cmdAdd_Click()
    Call ChangeCatcher
    
    NewRec = True
    lblRecNum.Visible = False
    lblCourseB.Caption = ""
    cmdSave.Default = True
    
    
    gPersonB.Course = ""
    
    txtCourse.Text = ""
    
    txtCourse.Tag = ""
    
    txtCourse.SetFocus
End Sub

Private Sub cmdClear_Click()
    If lstCourse = "" Then
        MsgBox "There are no records to delete.", vbOKOnly + vbExclamation, "Delete"
        Exit Sub
    End If
    
    Msg = "Are you sure you want to delete all records?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbYes Then
        If MySys.FileExists(App.Path + "\VOTING SYSTEM\Course.dat") = True Then
            Close #gFileNumB
            Kill App.Path + "\VOTING SYSTEM\Course.dat"
            
            gFileNumB = FreeFile
            Open App.Path + "\VOTING SYSTEM\Course.dat" For Random As gFileNumB Len = gRecordLenB
    
            gLastRecordB = 1
            gCurrentRecordB = gLastRecordB
            

            Call ListOfCourses
              
            ShowCourse
            MsgBox "All records cleared!", vb0konly + vbInformation, "Delete"
            lstCourse = lblCourseB.Caption
            lstCourse.SetFocus
            If lblCourseB.Caption = "" Then
                NewFile = True
                lblRecNum.Visible = False
            Else
                NewRec = False
            End If
        End If
    End If
    
    lstCourse = lblCourseB.Caption
    lstCourse.SetFocus
End Sub

Private Sub cmdDel_Click()

    If lblCourseB = "" Then
        MsgBox "There are no records to delete.", vbOKOnly + vbExclamation, "Delete"
        Exit Sub
    End If
    
    Dim DirResult
    Dim TmpFileNum
    Dim TmpPerson As Courses
    Dim RecNum As Long
    Dim TmpRecNum As Long
    Dim Msg As String
    Msg = "Delete " + UCase(txtCourse.Text) + "?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbNo Then
        lstCourse = lblCourseB.Caption
        lstCourse.SetFocus
        Exit Sub
    End If
    
    
   
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\MyCourse.tmp") = True Then
        Kill App.Path + "\VOTING SYSTEM\MyCourse.tmp"
    End If
 
    TmpFileNum = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\MyCourse.tmp" For Random As TmpFileNum Len = gRecordLenB
        
    RecNum = 1
    TmpRecNum = 1
    Do While RecNum < gLastRecordB + 1
        If RecNum <> gCurrentRecordB Then
            Get #gFileNumB, RecNum, TmpPerson
            Put #TmpFileNum, TmpRecNum, TmpPerson
            TmpRecNum = TmpRecNum + 1
        End If
        RecNum = RecNum + 1
    Loop
    
    Close #gFileNumB
    
    
    MySys.DeleteFile App.Path + "\VOTING SYSTEM\Course.dat"
    
    Close #TmpFileNum
    Name App.Path + "\VOTING SYSTEM\MyCourse.tmp" As App.Path + "\VOTING SYSTEM\Course.dat"
    
    gFileNumB = FreeFile
    Open App.Path + "\VOTING SYSTEM\Course.dat" For Random As gFileNumB Len = gRecordLenB
    
    gLastRecordB = gLastRecordB - 1
    If gLastRecordB = 0 Then gLastRecordB = 1
    
    If gCurrentRecordB > gLastRecordB Then
        gCurrentRecordB = gLastRecordB
    End If
    
    MsgBox "Course deleted.", vb0konly + vbInformation, "Delete"
    Call ListOfCourses
    
    ShowCourse
    lstCourse = lblCourseB.Caption
    lstCourse.SetFocus
    If lblCourseB.Caption = "" Then
        NewFile = True
        lblRecNum.Visible = False
    Else
        NewRec = False
    End If
    End Sub

Private Sub cmdEdit_Click()
    txtCourse.SetFocus
    cmdSave.Default = True
    
End Sub

Private Sub cmdEnterVoterStats_Click()
    ChosenCourse = lblCourseB.Caption
    ListOfRecords
    If gCurrentRecord = 0 Then
        MsgBox "No one in " & ChosenCourse & " has registered to vote.", vbOKOnly + vbInformation, "Statistics"
        lstCourse.SetFocus
        lstCourse = lblCourseB.Caption
        Exit Sub
    End If
    fmeStart.Visible = 0
    ShowCurrentRecord
    lstRecords.SetFocus
    lstRecords = lblName.Caption
End Sub

Private Sub cmdExit_Click()
    fmeStart.Visible = True
    lstCourse.SetFocus
End Sub

Private Sub cmdPrint_Click()
'to print results
    Printer.Font = "courier new"
    Printer.Print
    Printer.Print
    Printer.Print Tab(10); "V O T E R   S T A T I S T I C S"
    Printer.Print
    Printer.Print Tab(2); "Name: " + lblName.Caption
    Printer.Print Tab(2); "Stud Num: " + lblStudNum.Caption
    Printer.Print Tab(2); "Course: " + lblCourse.Caption
    Printer.Print
    Printer.Print
    Printer.Print txtVotes.Text
    Printer.Print
    Printer.Print
    Printer.Print Tab(2); lblTimeTrack.Caption
    Printer.Print Tab(2); lblVoteTime.Caption
    
    Printer.EndDoc
End Sub

Private Sub cmdSave_Click()
    If txtCourse.Text = "" Then
        MsgBox "Fill up entry completely!", vbOKOnly + vbExclamation, "Error Save"
        txtCourse.SetFocus
        Exit Sub
    End If
    If NewRec = True And NewFile = False Then
            gLastRecordB = gLastRecordB + 1
            Put #gFileNumB, gLastRecordB, gPersonB
            gCurrentRecordB = gLastRecordB
            SaveCourse
            ShowCourse
            cmdEnterVoterStats.Default = True
    Else
            SaveCourse
            ShowCourse
            cmdEnterVoterStats.Default = True
            NewFile = False
    End If
    On Error GoTo 100
    lstCourse = lblCourseB.Caption
    lstCourse.SetFocus
    Exit Sub
100
    lstCourse = lblCourseB.Caption
    lstCourse.SetFocus
End Sub

Private Sub Command1_Click()
    frmAdministrator.Show
    Unload Me
End Sub

Private Sub Form_Load()
    'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'to open sa VOTER DATABASE
    gRecordLen = Len(gPerson)
    gFileNum = FreeFile
    Open App.Path + "\VOTING SYSTEM\Voters.dat" For Random As gFileNum Len = gRecordLen
    gCurrentRecord = 1
    gLastRecord = FileLen(App.Path + "\VOTING SYSTEM\Voters.dat") / gRecordLen
    If gLastRecord = 0 Then
        gLastRecord = 1
    End If
    
    
'to determine if there are no currently saved courses
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\Course.dat") = True Then
        NewFile = False
    Else
        NewFile = True
        MsgBox "No courses are currently saved!", vbOKOnly + vbInformation, "Course Entry"
    End If
    

'to open the COURSE DATABASE
    gRecordLenB = Len(gPersonB)

    gFileNumB = FreeFile

    Open App.Path + "\VOTING SYSTEM\Course.dat" For Random As gFileNumB Len = gRecordLenB

    gCurrentRecordB = 1
    gLastRecordB = FileLen(App.Path + "\VOTING SYSTEM\Course.dat") / gRecordLenB
    
    If gLastRecordB = 0 Then
        gLastRecordB = 1
        NewFile = True
    End If
    
    Changes = False
    Call ListOfCourses
    ShowCourse
    lstCourse = lblCourseB.Caption
    
End Sub

Private Sub ListOfRecords()
    'to show the registered voters (according to chosen course) in the lstRecords listbox
    Dim hold As Integer
    lstRecords.Clear
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, gPerson
        If Trim(gPerson.VCourse) = ChosenCourse Then
            lstRecords.AddItem Trim(gPerson.VFullName)
            hold = RecNum
        End If
    Next
    gCurrentRecord = hold
    
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #gFileNum
    Close #gFileNumB
End Sub



Private Sub lstCourse_Click()
    ChangeCatcher
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As Courses
    
    NameToSearch = lstCourse
    If NameToSearch = "" Then
        Exit Sub
    End If
    
    NameToSearch = UCase(NameToSearch)
    Found = False
    For RecNum = 1 To gLastRecordB
        Get #gFileNumB, RecNum, TmpPerson
        If NameToSearch = UCase(Trim(TmpPerson.Course)) Then
            Found = True
            Exit For
        End If
    Next
    If Found = True Then
        gCurrentRecordB = RecNum
        ShowCourse
        
    Else
        MsgBox "Course " + NameToSearch + " not found.", vbOKOnly + vbExclamation, "Search"
    End If
    NewRec = False
End Sub

Private Sub lstCourse_DblClick()
    Call cmdEnterVoterStats_Click
    
End Sub

Private Sub lstRecords_Click()
    'to show profile of selected (in lstRecords listbox) registered voter
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As VoterInfo
    
    NameToSearch = lstRecords
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
        ShowCurrentRecord
    Else
        MsgBox "Name " + NameToSearch + " not found.", vbOKOnly + vbExclamation, "Search"
    End If
End Sub

Private Sub ShowCurrentRecord()
    'to displayed registered voter
    Get #gFileNum, gCurrentRecord, gPerson
    
    lblName.Caption = Trim(gPerson.VFullName)
    lblStudNum.Caption = Trim(gPerson.VStudNum)
    lblCourse.Caption = ChosenCourse
    
    txtVotes.Text = Trim(gPerson.VVote)
    If txtVotes.Text = "" Then
        txtVotes.Text = vbNewLine + "   HAS NOT VOTED YET."
    End If
    
    lblTimeTrack.Caption = "Registered on " & Trim(gPerson.VTimeTrack)
    lblVoteTime.Caption = "Voted on " & Trim(gPerson.VVoteTimeTrack)
    
    lblRecordNum.Caption = "Voter #" + _
                    Str(gCurrentRecord) + " of " + _
                    Str(gLastRecord) + " registered voter(s)."
    
    Exit Sub
End Sub

Private Sub lstRecords_GotFocus()
    Call lstRecords_Click
End Sub

Private Sub lstRecords_KeyUp(KeyCode As Integer, Shift As Integer)
    Call lstRecords_Click
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = "DATE: " & Date & "   TIME: " & Time
End Sub

Private Sub ListOfCourses()
    lstCourse.Clear
    For RecNumB = 1 To gLastRecordB
        Get #gFileNumB, RecNumB, gPersonB
        lstCourse.AddItem UCase(Trim(gPersonB.Course))
    Next
End Sub

Private Sub ShowCourse()
    Get #gFileNumB, gCurrentRecordB, gPersonB
    
    
    txtCourse.Text = Trim(gPersonB.Course)
    
    txtCourse.Tag = txtCourse.Text
    
    lblCourseB.Caption = UCase(txtCourse.Text)
    
    
    
    If lblCourseB.Caption <> "" Then
    lblRecNum.Visible = True
    End If
    lblRecNum.Caption = "Course #" + _
                    Str(gCurrentRecordB) + " of " + _
                    Str(gLastRecordB) + " saved course(s)."
    NewRec = False
                    
End Sub

Private Sub SaveCourse()
     
    gPersonB.Course = UCase(txtCourse.Text)
    
    Put #gFileNumB, gCurrentRecordB, gPersonB
    
    NewRec = False
    
    MsgBox "Course saved.", vbOKOnly + vbInformation, "Save"
    Call ListOfCourses
End Sub

Private Sub ChangeCatcher()
    If txtCourse.Tag <> txtCourse.Text Then
        GoTo SaveNow
    End If
    Exit Sub
SaveNow:
        If lblCourseB.Caption = "" Then
            Msg = "Save changes to this new record?"
        Else
            Msg = "Save changes to the record of " + UCase(lblCourseB.Caption) + "?"
        End If
        Response = MsgBox(Msg, vbYesNo + vbQuestion, "Save" & Save)
        If Response = vbYes Then
            Call cmdSave_Click
        Else
            Exit Sub
        End If
End Sub
Private Sub txtCourse_GotFocus()
    txtCourse.SelStart = 0
    txtCourse.SelLength = Len(txtCourse.Text)
End Sub
