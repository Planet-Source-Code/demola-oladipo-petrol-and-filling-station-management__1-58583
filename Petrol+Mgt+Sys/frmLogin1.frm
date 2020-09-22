VERSION 5.00
Begin VB.Form frmLogin1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Administrator Password"
   ClientHeight    =   2085
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4050
   Icon            =   "frmLogin1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1231.887
   ScaleMode       =   0  'User
   ScaleWidth      =   3802.731
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1545
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   960
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1500
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1500
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Retype Password:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   975
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Current Password:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      Caption         =   "&New Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtPassword = "" Then
        MsgBox "Please enter a New Password Value", vbCritical
        Text1.SetFocus
        SendKeys "{home}+{End}"
        Exit Sub
    End If
    If txtPassword <> Text1 Then
        MsgBox "The two entries of the new password do not match!", vbCritical
        Text1.SetFocus
        SendKeys "{home}+{End}"
        Exit Sub
    End If
    
    Dim sSQL As String
    sSQL = "select tblPassword from tblSecure where tblUsername = 'Administrator'"
    Set rsLogin = cn.Execute(sSQL)
    On Error GoTo Ouch
    If LCase$(txtUserName) = LCase$(rsLogin.Fields(0)) Then
        sSQL = "update tblSecure set tblPassword = '" & txtPassword & "' where tblUsername = 'Administrator'"
        cn.BeginTrans
        
        cn.Execute sSQL
        cn.CommitTrans
    End If
    MsgBox "Password Successfully Changed!", vbInformation
    Unload Me
    
    Exit Sub
Ouch:
    cn.RollbackTrans
End Sub

Private Sub Form_Load()
    'disableMainMenu False
    frmMain.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'disableMainMenu True
    frmMain.Enabled = True
End Sub

