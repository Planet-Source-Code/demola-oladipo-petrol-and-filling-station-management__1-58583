VERSION 5.00
Begin VB.Form frmPumpMgt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pumps Management"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmPumpMgt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4920
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Pump"
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit Pump Details"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Pump"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtReset 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Highest Value"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Product"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pump List"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   3375
         ItemData        =   "frmPumpMgt.frx":030A
         Left            =   120
         List            =   "frmPumpMgt.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmPumpMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Ouch
    If Command1.Caption <> "Save" Then
        MsgBox "Please select a Product Type", vbInformation
        Frame2.Enabled = True
        Command1.Caption = "Save"
        Command2.Enabled = False
        Command3.Enabled = False
    Else
        If Combo1.Text = "" Then
            MsgBox "Please select a Product Type", vbCritical
            Exit Sub
        End If
        If txtReset = "" Then
            MsgBox "Please enter the Highest Value by which the Pump will be reset!", vbExclamation
            Exit Sub
        End If
        
        On Error GoTo dError
        Dim dReset As Double
        dReset = CDbl(txtReset)
        
        On Error GoTo Ouch
        Frame2.Enabled = False
        
        Dim i As Integer
EnterID:
        i = CInt(InputBox("Please enter a number for the New Pump"))
        If i = 0 Then
            MsgBox "Invalid Pump ID", vbCritical
            Exit Sub
        End If
        sSQL = "insert into tblPumpType(tblPumpID, tblPumpType, tblLastReading, tblResetValue) values (" & i & ",'" & Combo1.Text & "', 0, " & txtReset & ")"
        cn.Execute sSQL
        
        loadPumps
        Command2.Enabled = True
        Command3.Enabled = True
        Command1.Caption = "New Pump"
    End If
    Exit Sub
Ouch:
    If Err.Number = -2147467259 Then
        MsgBox "You can not have duplicate Pump IDs", vbCritical
        Exit Sub
    Else
        MsgBox Err.Number & Err.Description
    End If
    
    Exit Sub
    
dError:
    MsgBox "Please enter a numeric value for the highest reset value of the pump", vbExclamation
    Exit Sub
    
End Sub

Private Sub Command2_Click()
    If Command2.Caption <> "Save" Then
        If List1.ListIndex < 0 Then
            MsgBox "Please select a pump to edit", vbInformation
            Exit Sub
        End If
        Frame2.Enabled = True
        Command2.Caption = "Save"
        Command1.Enabled = False
        Command3.Enabled = False
    Else
       If Combo1.Text = "" Then
            MsgBox "Please select a Product Type", vbCritical
            Exit Sub
        End If
        If MsgBox("are you sure you want to edit this pump?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        Frame2.Enabled = False
        Command2.Caption = "Edit Pump Details"
        sSQL = "update tblPumpType set tblPumpType= '" & Combo1.Text & "' where tblPumpId = " & List1.Text
        cn.Execute sSQL
        
        sSQL = "update tblPumpType set tblResetValue = " & txtReset & " where tblPumpId = " & List1.Text
        cn.Execute sSQL
        
        SaveStringSetting strProdName, "Config", "Pump Management", "True"
        loadPumps
        Command1.Enabled = True
        Command3.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
    If List1.ListIndex < 0 Then
        MsgBox "Please select a pump to Delete", vbInformation
        Exit Sub
    End If
    If MsgBox("are you sure you want to delete this pump?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    sSQL = "delete from tblPumpType where tblPumpID = " & List1.Text
    cn.Execute sSQL
    Combo1.Text = ""
    Frame2.Enabled = False
    loadPumps
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
    
    loadPumps
    
    sSQL = "select tblProductType from tblProductType"
    Set rsPumps = cn.Execute(sSQL)
    Combo1.Clear
    With rsPumps
        Do While Not .EOF
            Combo1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
    Frame2.Enabled = False
End Sub

Private Sub List1_Click()
    sSQL = "Select tblPumpType, tblResetValue from tblPumpType where tblPumpID = " & List1.Text
    Set rsPumps = cn.Execute(sSQL)
    Combo1.Text = rsPumps.Fields(0)
    txtReset.Text = rsPumps.Fields(1)
End Sub

Sub loadPumps()
    sSQL = "select tblPumpID from tblPumpType"
    Set rsPumps = cn.Execute(sSQL)
    List1.Clear
    With rsPumps
        Do While Not .EOF
            List1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With

End Sub
