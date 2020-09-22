VERSION 5.00
Begin VB.Form frmBonus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Staff Bonus - Female"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "frmBonus1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Target Sales - Bulk Oil"
      Height          =   1815
      Index           =   3
      Left            =   3600
      TabIndex        =   35
      Top             =   1200
      Width           =   3375
      Begin VB.TextBox txtNight 
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtMorning 
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNightTarget 
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtMorningTarget 
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "N"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   43
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "N"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Night"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   41
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Morning"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Target Sales - DPK"
      Height          =   1815
      Index           =   2
      Left            =   3600
      TabIndex        =   26
      Top             =   3120
      Width           =   3375
      Begin VB.TextBox txtMorningTarget 
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtNightTarget 
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtMorning 
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNight 
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Morning"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Night"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "N"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "N"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   31
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Target Sales - AGO"
      Height          =   1815
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   3375
      Begin VB.TextBox txtMorningTarget 
         Height          =   285
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtNightTarget 
         Height          =   285
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtMorning 
         Height          =   285
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNight 
         Height          =   285
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Morning"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Night"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "N"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "N"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unlock"
      Height          =   1695
      Left            =   3000
      TabIndex        =   16
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   735
      Left            =   5040
      TabIndex        =   15
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Target Sales - PMS"
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3375
      Begin VB.TextBox txtNight 
         Height          =   285
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtMorning 
         Height          =   285
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNightTarget 
         Height          =   285
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtMorningTarget 
         Height          =   285
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "N"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "N"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Night"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Morning"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Changes"
      Height          =   855
      Left            =   5040
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attendance"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtAbsent 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtPresent 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "If Absent"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "If Present"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Ouch
    
    Dim dblTemp As Double
    
    dblTemp = CDbl(txtPresent)
    dblTemp = CDbl(txtAbsent)
    For i = 0 To 3
        dblTemp = CDbl(txtMorningTarget(i))
        dblTemp = CDbl(txtNightTarget(i))
        dblTemp = CDbl(txtMorning(i))
        dblTemp = CDbl(txtNight(i))
    Next
    
    SaveStringSetting strProdName, "Options-Female", "Present", txtPresent
    SaveStringSetting strProdName, "Options-Female", "Absent", txtAbsent
        
    SaveStringSetting strProdName, "Options-Female", "Morning Target0", txtMorningTarget(0)
    SaveStringSetting strProdName, "Options-Female", "Night Target0", txtNightTarget(0)
    SaveStringSetting strProdName, "Options-Female", "Morning Incentive0", txtMorning(0)
    SaveStringSetting strProdName, "Options-Female", "Night Incentive0", txtNight(0)
    
    SaveStringSetting strProdName, "Options-Female", "Morning Target1", txtMorningTarget(1)
    SaveStringSetting strProdName, "Options-Female", "Night Target1", txtNightTarget(1)
    SaveStringSetting strProdName, "Options-Female", "Morning Incentive1", txtMorning(1)
    SaveStringSetting strProdName, "Options-Female", "Night Incentive1", txtNight(1)
    
    SaveStringSetting strProdName, "Options-Female", "Morning Target2", txtMorningTarget(2)
    SaveStringSetting strProdName, "Options-Female", "Night Target2", txtNightTarget(2)
    SaveStringSetting strProdName, "Options-Female", "Morning Incentive2", txtMorning(2)
    SaveStringSetting strProdName, "Options-Female", "Night Incentive2", txtNight(2)
    
    SaveStringSetting strProdName, "Options-Female", "Morning Target3", txtMorningTarget(3)
    SaveStringSetting strProdName, "Options-Female", "Night Target3", txtNightTarget(3)
    SaveStringSetting strProdName, "Options-Female", "Morning Incentive3", txtMorning(3)
    SaveStringSetting strProdName, "Options-Female", "Night Incentive3", txtNight(3)
    
    SaveStringSetting strProdName, "Config", "Female Bonus", "True"
    
    txtAbsent.Locked = True
    txtPresent.Locked = True
    For i = 0 To 3
        txtMorning(i).Locked = True
        txtMorningTarget(i).Locked = True
        txtNight(i).Locked = True
        txtNightTarget(i).Locked = True
    Next
    Exit Sub
Ouch:
    MsgBox "There is an error in your values. Please check!", vbCritical

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    'display login form
    frmLogin.Show 1
    If LoginSucceeded Then
        Unload frmLogin
        txtAbsent.Locked = False
        txtPresent.Locked = False
        For i = 0 To 3
            txtMorning(i).Locked = False
            txtMorningTarget(i).Locked = False
            txtNight(i).Locked = False
            txtNightTarget(i).Locked = False
        Next
        LoginSucceeded = False
    End If
End Sub

Private Sub Form_Activate()
Me.Enabled = False
    'display login form
    frmLogin.Show 1
    If LoginSucceeded Then
        Unload frmLogin
        Me.Enabled = True
        LoginSucceeded = False
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtPresent = GetStringSetting(strProdName, "Options-Female", "Present", "100")
    txtAbsent = GetStringSetting(strProdName, "Options-Female", "Absent", "-50")
    
    txtMorningTarget(0) = GetStringSetting(strProdName, "Options-Female", "Morning Target0", "3300")
    txtNightTarget(0) = GetStringSetting(strProdName, "Options-Female", "Night Target0", "4000")
    txtMorning(0) = GetStringSetting(strProdName, "Options-Female", "Morning Incentive0", "100")
    txtNight(0) = GetStringSetting(strProdName, "Options-Female", "Night Incentive0", "100")
    
    txtMorningTarget(1) = GetStringSetting(strProdName, "Options-Female", "Morning Target1", "3300")
    txtNightTarget(1) = GetStringSetting(strProdName, "Options-Female", "Night Target1", "4000")
    txtMorning(1) = GetStringSetting(strProdName, "Options-Female", "Morning Incentive1", "100")
    txtNight(1) = GetStringSetting(strProdName, "Options-Female", "Night Incentive1", "100")
    
    txtMorningTarget(2) = GetStringSetting(strProdName, "Options-Female", "Morning Target2", "3300")
    txtNightTarget(2) = GetStringSetting(strProdName, "Options-Female", "Night Target2", "4000")
    txtMorning(2) = GetStringSetting(strProdName, "Options-Female", "Morning Incentive2", "100")
    txtNight(2) = GetStringSetting(strProdName, "Options-Female", "Night Incentive2", "100")
    
    txtMorningTarget(3) = GetStringSetting(strProdName, "Options-Female", "Morning Target3", "3300")
    txtNightTarget(3) = GetStringSetting(strProdName, "Options-Female", "Night Target3", "4000")
    txtMorning(3) = GetStringSetting(strProdName, "Options-Female", "Morning Incentive3", "100")
    txtNight(3) = GetStringSetting(strProdName, "Options-Female", "Night Incentive3", "100")
End Sub
