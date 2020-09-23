VERSION 5.00
Begin VB.Form frmPassAdmin 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4695
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmPassAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000C&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H8000000C&
      Caption         =   "&Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Enter Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Enter Password:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1365
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuNew 
         Caption         =   "Set &New Password"
      End
   End
End
Attribute VB_Name = "frmPassAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MySys As New FileSystemObject
Private Sub cmdCancel_Click()
    Select Case ToPass
        Case 0
            Unload Me
        Case 1
            frmMain.Enabled = True
        Case 2
            frmVote.Enabled = True
        Case 3
            frmVoterRegistration.Enabled = True
        Case 4
            frmAdministrator.Enabled = True
    End Select
    ToPass = 0
    Unload Me
End Sub

Private Sub cmdEnter_Click()
    
    Open "c:\WinSysVsp.dat" For Input As #4
    Input #4, Password
    Close #4
    
    If txtPass.Text = Password Then
        MsgBox "ACCESS GRANTED", vbOKOnly + vbExclamation, "Password"
        Select Case ToPass
            Case 0
                frmMain.Show
                Unload Me
            Case 1
                frmAdministrator.Show
                Unload frmMain
                ToPass = 0
                Unload Me
            Case 2
                frmMain.Show
                Unload frmVote
                ToPass = 0
                Unload Me
            Case 3
                frmMain.Show
                Unload frmVoterRegistration
                ToPass = 0
                Unload Me
            Case 4
                ToPass = 0
                MySys.DeleteFolder App.Path + "\VOTING SYSTEM"
                MkDir App.Path + "\VOTING SYSTEM"
                MsgBox "All records deleted", vbOKOnly + vbExclamation, "Clear Records"
                frmAdministrator.Enabled = True
                Unload Me
            End Select
    Else
        MsgBox "ACCESS DENIED", vbOKOnly + vbExclamation, "Password"
        txtPass.Text = ""
        txtPass.SetFocus
        
    End If
    

End Sub

Private Sub Form_Load()
    Select Case ToPass
            Case 0
                lblDisplay.Caption = "VOTING SYSTEM > > > "
            Case 1
                lblDisplay.Caption = "You are trying to enter Administrator Access."
            Case 2
                lblDisplay.Caption = "You are trying to exit Election Process."
            Case 3
                lblDisplay.Caption = "You are trying to exit Voter Registration."
            Case 4
                lblDisplay.Caption = "WARNING!  You are trying to DELETE ALL RECORDS."
    End Select
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
On Error GoTo 100
    Open "c:\WinSysVsp.dat" For Input As #9
    Input #9, Password
    Close #9
    Exit Sub
100
    Open "c:\WinSysVsp.dat" For Output As #10
    Write #10, "daystar"
    Close #10
    
    Open "c:\WinSysVsp.dat" For Input As #11
    Input #11, Password
    Close #11
    
End Sub




Private Sub mnuSetPass_Click()
    frmSetPass.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #4
    Close #9
    Close #10
    Close #11
End Sub

Private Sub mnuNew_Click()
    frmSetPass.Show
End Sub
