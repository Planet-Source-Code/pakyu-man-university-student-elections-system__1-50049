VERSION 5.00
Begin VB.Form frmSetPass 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set New Password"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSetPassVotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4680
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H8000000C&
      Caption         =   "&Set"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtVerifyPass 
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
      Left            =   2280
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtNewPass 
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
      Left            =   2280
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtCurrentPass 
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
      Left            =   2280
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Caption         =   "Passwords are case sensitive."
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
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Maximum of 13 characters."
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
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Set your new password."
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
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Verify New Password:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Enter New Password:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Enter Current Password:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmSetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewPass As String


Private Sub cmdCancel_Click()
    frmPassAdmin.Show
    Unload Me
End Sub

Private Sub cmdSet_Click()
    Open "c:\WinSysVsp.dat" For Input As #1
    Input #1, Password
    Close #1
    
    
    If txtCurrentPass.Text <> Password Then
        MsgBox "You entered a wrong Current Password.", vbOKOnly + vbExclamation, "Password"
        txtCurrentPass.Text = ""
        txtCurrentPass.SetFocus
     
    ElseIf txtNewPass.Text <> txtVerifyPass.Text Then
        MsgBox "Your new password does not match with its verification.", vbOKOnly + vbExclamation, "Password"
        txtNewPass.Text = ""
        txtVerifyPass.Text = ""
        txtNewPass.SetFocus
    ElseIf txtNewPass.Text = "" Or txtVerifyPass.Text = "" And txtCurrentPass <> "" Then
        MsgBox "Enter for a New Password and its Verification.", vbOKOnly + vbExclamation, "Password"
        txtNewPass.Text = ""
        txtVerifyPass.Text = ""
        txtNewPass.SetFocus
    Else
        NewPass = txtVerifyPass.Text
        Open "c:\WinSysVsp.dat" For Output As #1
        Write #1, NewPass
        Close #1
        MsgBox "New password saved.", vbOKOnly + vbExclamation, "Password"
        frmPassAdmin.Show
        Unload Me
        
    End If
    
End Sub

Private Sub Form_Load()
    Unload frmPassAdmin
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
On Error GoTo 100
    Open "c:\WinSysVsp.dat" For Input As #1
    Input #1, Password
    Close #1
    Exit Sub
100
    Open "c:\WinSysVsp.dat" For Output As #1
    Write #1, "daystar"
    Close #1
    
    Open "c:\WinSysVsp.dat" For Input As #1
    Input #1, Password
    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPassAdmin.Show
    Unload Me
        
End Sub

Private Sub txtCurrentPass_GotFocus()
    txtCurrentPass.SelStart = 0
    txtCurrentPass.SelLength = Len(txtCurrentPass.Text)
    
End Sub

Private Sub txtNewPass_GotFocus()
    txtNewPass.SelStart = 0
    txtNewPass.SelLength = Len(txtNewPass.Text)
End Sub

Private Sub txtVerifyPass_GotFocus()
    txtVerifyPass.SelStart = 0
    txtVerifyPass.SelLength = Len(txtVerifyPass.Text)
End Sub
