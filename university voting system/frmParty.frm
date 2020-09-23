VERSION 5.00
Begin VB.Form frmParty 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Candidate Positions"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   6975
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
         TabIndex        =   16
         Top             =   240
         Width           =   60
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      Caption         =   "Current Parties Saved"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   6975
      Begin VB.ListBox lstRecords 
         Height          =   2595
         ItemData        =   "frmParty.frx":0000
         Left            =   120
         List            =   "frmParty.frx":0002
         TabIndex        =   0
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.TextBox txtParty 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   6975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   6015
      Left            =   7200
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
         Picture         =   "frmParty.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
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
         Picture         =   "frmParty.frx":0446
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
         MaskColor       =   &H8000000A&
         Picture         =   "frmParty.frx":0AB0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDel 
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
         Picture         =   "frmParty.frx":111A
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
         Picture         =   "frmParty.frx":121C
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
         Picture         =   "frmParty.frx":165E
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
      Width           =   6975
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "PARTY ENTRY"
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
         Width           =   2505
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
      Top             =   6120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label label1 
      BackColor       =   &H80000007&
      Caption         =   "Party:"
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
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gPerson As Parties
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
        If MySys.FileExists(App.Path + "\VOTING SYSTEM\Party.dat") = True Then
            Close #gFileNum
            Kill App.Path + "\VOTING SYSTEM\Party.dat"
            
            gFileNum = FreeFile
            Open App.Path + "\VOTING SYSTEM\Party.dat" For Random As gFileNum Len = gRecordLen
    
            gLastRecord = 1
            gCurrentRecord = gLastRecord
            

            Call ListOfParty
              
            ShowParty
            MsgBox "All records cleared!", vb0konly + vbInformation, "Delete"
            lstRecords = lblParty.Caption
            lstRecords.SetFocus
            If lblParty.Caption = "" Then
                NewFile = True
                lblRecNum.Visible = False
            Else
                NewRec = False
            End If
        End If
    End If
    
    lstRecords = lblParty.Caption
    lstRecords.SetFocus
End Sub

Private Sub cmdDel_Click()
    If lblParty = "" Then
        MsgBox "There are no records to delete.", vbOKOnly + vbExclamation, "Delete"
        Exit Sub
    End If
    
    Dim DirResult
    Dim TmpFileNum
    Dim TmpPerson As Parties
    Dim RecNum As Long
    Dim TmpRecNum As Long
    Dim Msg As String
    Msg = "Delete " + UCase(txtParty.Text) + "?"
    
    If MsgBox(Msg, vbYesNo + vbQuestion, "Delete") = vbNo Then
        lstRecords = lblParty.Caption
        lstRecords.SetFocus
        Exit Sub
    End If
    
    
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\MyParty.tmp") = True Then
        Kill App.Path + "\VOTING SYSTEM\MyParty.tmp"
    End If
    TmpFileNum = FreeFile
    
    Open App.Path + "\VOTING SYSTEM\MyParty.tmp" For Random As TmpFileNum Len = gRecordLen
        
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
    
    MySys.DeleteFile App.Path + "\VOTING SYSTEM\Party.dat"
    
    Close #TmpFileNum
    Name App.Path + "\VOTING SYSTEM\MyParty.tmp" As App.Path + "\VOTING SYSTEM\Party.dat"
    
    gFileNum = FreeFile
    Open App.Path + "\VOTING SYSTEM\Party.dat" For Random As gFileNum Len = gRecordLen
    
    gLastRecord = gLastRecord - 1
    If gLastRecord = 0 Then gLastRecord = 1
    
    If gCurrentRecord > gLastRecord Then
        gCurrentRecord = gLastRecord
    End If
    
    MsgBox "Party deleted.", vb0konly + vbInformation, "Delete"
    Call ListOfParty
    
    ShowParty
    lstRecords = lblParty.Caption
    lstRecords.SetFocus
    If lblParty.Caption = "" Then
        NewFile = True
        lblRecNum.Visible = False
    Else
        NewRec = False
    End If
    
End Sub

Private Sub cmdEdit_Click()
    txtParty.SetFocus
End Sub

Private Sub cmdNew_Click()
    Call ChangeCatcher
    
    NewRec = True
    lblRecNum.Visible = False
    lblParty.Caption = ""
    
    gPerson.Party = ""
    txtParty.Text = ""
    txtParty.Tag = ""
    txtParty.SetFocus
    
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
        ShowParty
    End If
    lstRecords = lblParty.Caption
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
        ShowParty
    End If
    lstRecords = lblParty.Caption
    lstRecords.SetFocus
End Sub

Private Sub cmdSave_Click()
    If txtParty.Text = "" Then
        MsgBox "Fill up entry completely!", vbOKOnly + vbExclamation, "Error Save"
        txtParty.SetFocus
        Exit Sub
    End If
    
    If NewRec = True And NewFile = False Then
            gLastRecord = gLastRecord + 1
            Put #gFileNum, gLastRecord, gPerson
            gCurrentRecord = gLastRecord
            SaveParty
            ShowParty
    Else
            SaveParty
            ShowParty
            NewFile = False
    End If
    On Error GoTo 100
    lstRecords = lblParty.Caption
    lstRecords.SetFocus
    Exit Sub
100
    lstRecords = lblParty.Caption
    lstRecords.SetFocus
End Sub


Private Sub Form_Load()
'to center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

'to determine if there are no currently saved party
    If MySys.FileExists(App.Path + "\VOTING SYSTEM\Party.dat") = True Then
        NewFile = False
    Else
        NewFile = True
        MsgBox "No parties are currently saved!", vbOKOnly + vbInformation, "Party Entry"
    End If
   

'to open the Party DATABASE
    gRecordLen = Len(gPerson)

    gFileNum = FreeFile

    Open App.Path + "\VOTING SYSTEM\Party.dat" For Random As gFileNum Len = gRecordLen

    gCurrentRecord = 1
    gLastRecord = FileLen(App.Path + "\VOTING SYSTEM\Party.dat") / gRecordLen
    
    If gLastRecord = 0 Then
        gLastRecord = 1
        NewFile = True
    End If
    
    
    Call ListOfParty
    ShowParty
    lstRecords = lblParty.Caption
    
End Sub

Private Sub ShowParty()
    Get #gFileNum, gCurrentRecord, gPerson
    
    txtParty.Text = Trim(gPerson.Party)
    txtParty.Tag = txtParty.Text
    lblParty.Caption = UCase(txtParty.Text)
    
    If lblParty.Caption <> "" Then
        lblRecNum.Visible = True
    End If
    lblRecNum.Caption = "Party #" + _
                    Str(gCurrentRecord) + " of " + _
                    Str(gLastRecord) + " saved party(s)."
    NewRec = False
End Sub

Private Sub SaveParty()
    
    gPerson.Party = UCase(txtParty.Text)
    
    Put #gFileNum, gCurrentRecord, gPerson
    
    NewRec = False
    
    MsgBox "Party saved.", vbOKOnly + vbInformation, "Save"
    Call ListOfParty
End Sub

Private Sub ChangeCatcher()
    If txtParty.Tag <> txtParty.Text Then
        GoTo SaveNow
    End If
    Exit Sub
SaveNow:
        If lblParty.Caption = "" Then
            Msg = "Save changes to this new record?"
        Else
            Msg = "Save changes to the record of " + UCase(lblParty.Caption) + "?"
        End If
        Response = MsgBox(Msg, vbYesNo + vbQuestion, "Save" & Save)
        If Response = vbYes Then
            Call cmdSave_Click
        Else
            Exit Sub
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Close #gFileNum
End Sub

Private Sub ListOfParty()
    lstRecords.Clear
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, gPerson
        lstRecords.AddItem UCase(Trim(gPerson.Party))
    Next
End Sub





Private Sub lstRecords_Click()
    ChangeCatcher
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpPerson As Parties
    
    NameToSearch = lstRecords
    If NameToSearch = "" Then
        Exit Sub
    End If
    
    NameToSearch = UCase(NameToSearch)
    Found = False
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, TmpPerson
        If NameToSearch = UCase(Trim(TmpPerson.Party)) Then
            Found = True
            Exit For
        End If
    Next
    If Found = True Then
        gCurrentRecord = RecNum
        ShowParty
    Else
        MsgBox "Party " + NameToSearch + " not found.", vbOKOnly + vbExclamation, "Search"
    End If
    
End Sub

Private Sub txtParty_GotFocus()
    txtParty.SelStart = 0
    txtParty.SelLength = Len(txtParty.Text)
End Sub
