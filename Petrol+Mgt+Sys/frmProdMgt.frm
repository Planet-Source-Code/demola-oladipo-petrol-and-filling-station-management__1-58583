VERSION 5.00
Begin VB.Form frmProdMgt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Management"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmProdMgt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBULK 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unlock"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtPMS 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtDPK 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtAGO 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Unit Cost per Litre of BULK"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Unit Cost per Litre of PMS"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Unit Cost per Litre of DPK"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Unit Cost per Litre of AGO"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmProdMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String

Private Sub Command1_Click()
    If txtAGO = "" Then
        MsgBox "please enter a value for AGO"
        txtAGO.SetFocus
        Exit Sub
    End If
    If txtDPK = "" Then
        MsgBox "please enter a value for DPK"
        txtDPK.SetFocus
        Exit Sub
    End If
    If txtPMS = "" Then
        MsgBox "please enter a value for PMS"
        txtPMS.SetFocus
        Exit Sub
    End If
    If txtBULK = "" Then
        MsgBox "please enter a value for BULK Oil"
        txtAGO.SetFocus
        Exit Sub
    End If

    SaveStringSetting strProdName, "Config", "Product Management", "True"
    If Command1.Caption <> "Close" Then
        sSQL = "update tblProductType set tblCost = " & CDbl(txtAGO) & " where tblProductType = 'AGO'"
        cn.Execute sSQL
        
        sSQL = "update tblProductType set tblCost = " & CDbl(txtDPK) & " where tblProductType = 'DPK'"
        cn.Execute sSQL
        
        sSQL = "update tblProductType set tblCost = " & CDbl(txtPMS) & " where tblProductType = 'PMS'"
        cn.Execute sSQL
        
        sSQL = "update tblProductType set tblCost = " & CDbl(txtBULK) & " where tblProductType = 'BULK'"
        cn.Execute sSQL
        
        Command1.Caption = "Close"
        txtAGO.Locked = True
        txtDPK.Locked = True
        txtPMS.Locked = True
        txtBULK.Locked = True
    Else
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    'display login form
    frmLogin.Show 1
    If LoginSucceeded Then
        Unload frmLogin
        txtAGO.Locked = False
        txtDPK.Locked = False
        txtPMS.Locked = False
        txtBULK.Locked = False
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
    
    sSQL = "select tblCost from tblProductType where tblProductType = 'AGO'"
    Set rs = cn.Execute(sSQL)
    txtAGO = rs.Fields(0)
    
    sSQL = "select tblCost from tblProductType where tblProductType = 'DPK'"
    Set rs = cn.Execute(sSQL)
    txtDPK = rs.Fields(0)
    
    sSQL = "select tblCost from tblProductType where tblProductType = 'PMS'"
    Set rs = cn.Execute(sSQL)
    txtPMS = rs.Fields(0)
    
    sSQL = "select tblCost from tblProductType where tblProductType = 'BULK'"
    Set rs = cn.Execute(sSQL)
    txtBULK = rs.Fields(0)
    
End Sub

Private Sub txtAGO_Change()
    Command1.Caption = "Save Changes"
End Sub

Private Sub txtDPK_Change()
    Command1.Caption = "Save Changes"
End Sub

Private Sub txtPMS_Change()
    Command1.Caption = "Save Changes"
End Sub

Private Sub txtbulk_Change()
    Command1.Caption = "Save Changes"
End Sub

