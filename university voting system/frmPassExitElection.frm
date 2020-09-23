VERSION 5.00
Begin VB.Form frmPassExitElection 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   2055
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   2400
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2400
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
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
      Top             =   1080
      Width           =   1935
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "You are trying to exit ELECTION!"
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   240
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmPassExitElection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
    frmVote.Enabled = True
    Unload Me
End Sub

Private Sub cmdEnter_Click()
    
    Open "c:\eep.dat" For Input As #7
    Input #7, Password
    Close #7
    
    If txtPass.Text = Password Then
        MsgBox "ACCESS GRANTED", vbOKOnly + vbExclamation, "Password"
        frmMain.Show
        Unload Me
        Unload frmVote
    Else
        MsgBox "ACCESS DENIED", vbOKOnly + vbExclamation, "Password"
        txtPass.Text = ""
        txtPass.SetFocus
        
    End If
    'MsgBox "shit"

End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
On Error GoTo 100
    Open "c:\eep.dat" For Input As #3
    Input #3, Password
    Close #3
    Exit Sub
100
    Open "c:\eep.dat" For Output As #6
    Write #6, "election"
    Close #6
    
    Open "c:\eep.dat" For Input As #8
    Input #8, Password
    Close #8
    
End Sub




Private Sub mnuSetPass_Click()
    frmSetPass.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #3
    Close #6
    Close #7
    Close #8
End Sub
