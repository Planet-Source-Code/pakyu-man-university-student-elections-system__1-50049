VERSION 5.00
Begin VB.Form frmPositions 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Candidate Positions"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   6255
      Begin VB.Label lblPos 
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
         TabIndex        =   16
         Top             =   240
         Width           =   60
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      Caption         =   "Current Positions Saved"
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
      Height          =   4575
      Left            =   3720
      TabIndex        =   14
      Top             =   1560
      Width           =   2655
      Begin VB.ListBox lstRecords 
         Height          =   4155
         ItemData        =   "frmPositions.frx":0000
         Left            =   120
         List            =   "frmPositions.frx":0002
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txtPosition 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   6015
      Left            =   6480
      TabIndex        =   12
      Top             =   120
      Width           =   1695
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
         TabIndex        =   6
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
         Picture         =   "frmPositions.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdSavePos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Save"
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
         Picture         =   "frmPositions.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdPosExit 
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
         TabIndex        =   9
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton cmdNewPos 
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
         MaskColor       =   &H8000000A&
         Picture         =   "frmPositions.frx":0AB0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelPos 
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
         Picture         =   "frmPositions.frx":111A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nex&t"
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
         Picture         =   "frmPositions.frx":121C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pre&vious"
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
         Picture         =   "frmPositions.frx":165E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4320
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6255
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "CANDIDATE POSITIONS ENTRY"
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
         TabIndex        =   11
         Top             =   240
         Width           =   5460
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
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label label1 
      BackColor       =   &H80000007&
      Caption         =   "Position:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   3015
   End
End
Attribute VB_Name = "frmPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gPerson As Positions
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim gCurrentRecord As Long
Dim gLastRecord As Long
Dim MySys As New FileSystemObject
Dim NewRec As Boolean
Dim NewFile As Boolean

Private Sub cmdClear_Click()
    If lstRecords = "" Then
        MsgBox "There are no records to delete.", vbOKOnly + vbExclamation, "Delete"
        Exit Sub
    End If
    
    Msg = "Are you sure you want ot delete all records?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbYes Then
        If MySys.FileExists(App.Path + "\VOTING SYSTEM\Positions.dat") = True Then
            Close #gFileNum
            Kill App.Path + "\VOTING SYSTEM\Positions.dat"
            
            gFileNum = FreeFile
            Open App.Path + "\VOTING SYSTEM\Positions.dat" For Random As gFileNum Len = gRecordLen
    
            gLastRecord = 1
            gCurrentRecord = gLastRecord

            Call ListOfPositions
              
            ShowPositions
            MsgBox "All records cleared!", vb0konly + vbInformation, "Delete"
            lstRecords = lblPos.Caption
            lstRecords.SetFocus
            If lblPos.Caption = "" Then
                NewFile = True
                lblRecNum.Visible = False
            Else
                NewRec = False
            End If
        End If
    End If
    
    lstRecords = lblPos.Caption
    lstRecords.SetFocus
End Sub

Private Sub cmdDelPos_Click()
    If lblPos = "" Then
        MsgBox "There are no records to delete.", vbOKOnly + vbExclamation, "Delete"
        Exit Sub
    End If
    
    Dim DirResult
    Dim TmpFileNum
    Dim TmpPerson As Positions
    Dim RecNum As Long
    Dim TmpRecNum As Long
    Dim Msg As String
    Msg = "Delete " + UCase(txtPosition.Text) + "?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbNo Then
        lstRecords = lblPos.Caption
        lstRecords.SetFocus
        Exit Sub
    End If
    
    
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\MyPos.tmp") = True Then
        Kill App.Path + "\VOTING SYSTEM\MyPos.tmp"
    End If
    TmpFileNum = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\MyPos.tmp" For Random As TmpFileNum Len = gRecordLen
        
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
    
    
    MySys.DeleteFile App.Path + "\VOTING SYSTEM\Positions.dat"
    
    Close #TmpFileNum
    Name App.Path + "\VOTING SYSTEM\MyPos.tmp" As App.Path + "\VOTING SYSTEM\Positions.dat"
    
    gFileNum = FreeFile
    Open App.Path + "\VOTING SYSTEM\Positions.dat" For Random As gFileNum Len = gRecordLen
    
    gLastRecord = gLastRecord - 1
    If gLastRecord = 0 Then gLastRecord = 1
    
    If gCurrentRecord > gLastRecord Then
        gCurrentRecord = gLastRecord
    End If
    
    MsgBox "Position deleted.", vb0konly + vbInformation, "Delete"
    Call ListOfPositions
    
    ShowPositions
    lstRecords = lblPos.Caption
    lstRecords.SetFocus
    If lblPos.Caption = "" Then
       NewFile = True
       lblRecNum.Visible = False
    Else
       NewRec = False
    End If
    
    
End Sub

Private Sub cmdEdit_Click()
    txtPosition.SetFocus
End Sub

Private Sub cmdNewPos_Click()
    Call ChangeCatcher
    
    NewRec = True
    lblRecNum.Visible = False
    lblPos.Caption = ""
    
    
    gPerson.Position = ""
    
    txtPosition.Text = ""
    
    txtPosition.Tag = ""
    
    txtPosition.SetFocus
End Sub

Private Sub cmdNext_Click()
    If cmdPrev.Enabled = 0 Then
        cmdPrev.Enabled = 1
    End If
    If gCurrentRecord = gLastRecord Then
        Beep
        cmdNext.Enabled = 0
        cmdPrev.SetFocus
    Else
        Call ChangeCatcher
        gCurrentRecord = gCurrentRecord + 1
        ShowPositions
    End If
    lstRecords = lblPos.Caption
    lstRecords.SetFocus
End Sub

Private Sub cmdPosExit_Click()
    ChangeCatcher
    frmAdministrator.Show
    Unload Me
End Sub

Private Sub cmdPrev_Click()
    If cmdNext.Enabled = 0 Then
        cmdNext.Enabled = 1
    End If
    
    If gCurrentRecord = 1 Then
        Beep
        cmdPrev.Enabled = 0
        cmdNext.SetFocus
    Else
        Call ChangeCatcher
        gCurrentRecord = gCurrentRecord - 1
        ShowPositions
    End If
    lstRecords = lblPos.Caption
    lstRecords.SetFocus
End Sub

Private Sub cmdSavePos_Click()
    If txtPosition.Text = "" Then
        MsgBox "Fill up entry completely!", vbOKOnly + vbExclamation, "Error Save"
        txtPosition.SetFocus
        Exit Sub
    End If
    If NewRec = True And NewFile = False Then
            gLastRecord = gLastRecord + 1
            Put #gFileNum, gLastRecord, gPerson
            gCurrentRecord = gLastRecord
            SavePositions
            ShowPositions
    Else
            SavePositions
            ShowPositions
            NewFile = False
    End If
    On Error GoTo 100
    lstRecords = lblPos.Caption
    lstRecords.SetFocus
    Exit Sub
100
    lstRecords = lblPos.Caption
    lstRecords.SetFocus
End Sub


Private Sub Form_Load()
'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

'to determine if there are no currently saved positions
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\Positions.dat") = True Then
        NewFile = False
    Else
        NewFile = True
        MsgBox "No candidate positions are currently saved!", vbOKOnly + vbInformation, "Candidates"
    End If
    
'to open the POSITION DATABASE
    gRecordLen = Len(gPerson)

    gFileNum = FreeFile

    Open App.Path + "\VOTING SYSTEM\Positions.dat" For Random As gFileNum Len = gRecordLen

    gCurrentRecord = 1
    gLastRecord = FileLen(App.Path + "\VOTING SYSTEM\Positions.dat") / gRecordLen
    
    If gLastRecord = 0 Then
        gLastRecord = 1
        NewFile = True
    End If
    
    Call ListOfPositions
    ShowPositions
    lstRecords = lblPos.Caption
    
End Sub

Private Sub ShowPositions()
    Get #gFileNum, gCurrentRecord, gPerson
    
    
    txtPosition.Text = Trim(gPerson.Position)
    
    txtPosition.Tag = txtPosition.Text
    
    lblPos.Caption = UCase(txtPosition.Text)
    
    
    
    If lblPos.Caption <> "" Then
    lblRecNum.Visible = True
    End If
    lblRecNum.Caption = "Position #" + _
                    Str(gCurrentRecord) + " of " + _
                    Str(gLastRecord) + " saved position(s)."
    NewRec = False
                    
End Sub

Private Sub SavePositions()
    
    gPerson.Position = UCase(txtPosition.Text)
    
    Put #gFileNum, gCurrentRecord, gPerson
    
    NewRec = False
    
    MsgBox "Position saved.", vbOKOnly + vbInformation, "Save"
    Call ListOfPositions
End Sub

Private Sub ChangeCatcher()
    If txtPosition.Tag <> txtPosition.Text Then
        GoTo SaveNow
    End If
    Exit Sub
SaveNow:
        If lblPos.Caption = "" Then
            Msg = "Save changes to this new record?"
        Else
            Msg = "Save changes to the record of " + UCase(lblPos.Caption) + "?"
        End If
        Response = MsgBox(Msg, vbYesNo + vbQuestion, "Save" & Save)
        If Response = vbYes Then
            Call cmdSavePos_Click
        Else
            Exit Sub
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Close #gFileNum
End Sub

Private Sub ListOfPositions()
    lstRecords.Clear
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, gPerson
        lstRecords.AddItem UCase(Trim(gPerson.Position))
    Next
End Sub

Private Sub lstRecords_Click()
    ChangeCatcher
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As Positions
    
    NameToSearch = lstRecords
    If NameToSearch = "" Then
        Exit Sub
    End If
    
    NameToSearch = UCase(NameToSearch)
    Found = False
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, TmpPerson
        If NameToSearch = UCase(Trim(TmpPerson.Position)) Then
            Found = True
            Exit For
        End If
    Next
    If Found = True Then
        gCurrentRecord = RecNum
        ShowPositions
    Else
        MsgBox "Position " + NameToSearch + " not found.", vbOKOnly + vbExclamation, "Search"
    End If
    
End Sub


Private Sub txtPosition_GotFocus()
    txtPosition.SelStart = 0
    txtPosition.SelLength = Len(txtPosition.Text)
End Sub
