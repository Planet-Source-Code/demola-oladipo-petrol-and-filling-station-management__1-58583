VERSION 5.00
Begin VB.Form frmPumpLitre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Litres Sold"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmPumpLitre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6600
   Begin VB.TextBox txtBackdate 
      Enabled         =   0   'False
      Height          =   405
      Left            =   4680
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Temporary Backdate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5160
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5400
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtReturn 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "&Compute"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1560
      TabIndex        =   18
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4800
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtAmount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """N""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox txtSold 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox txtWasted 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Quantity Return to Tank"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtPresentReading 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pump Details"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4455
      Begin VB.TextBox txtCost 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtLR 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtPT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Product Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Last Reading"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Product Type"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboPump 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label9 
      Caption         =   "Reason"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Quantity Returned"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Litres Sold"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Present Reading"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Select Pump"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPumpLitre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dblPR, dblWaste, dblLR, dblSold, dblTemp, dblCost, dblAmount As Double

Private Sub cboPump_Click()
    On Error Resume Next
    Dim sSQL As String
    sSQL = "select * from tblPumpType where tblPumpID=" & CInt(cboPump.Text)
    Set rs = cn.Execute(sSQL)
    txtPT = rs.Fields("tblPumpType")
    txtLR = rs.Fields("tblLastReading")
    Text2 = rs.Fields("tblResetValue")
    
    sSQL = "select tblCost from tblProductType where tblProductType = '" & txtPT & "'"
    Set rs = cn.Execute(sSQL)
    txtCost = rs.Fields(0)
    txtWasted.Enabled = True
    txtPresentReading.Enabled = True
    
    Check1.Enabled = True
    cmdCompute.Enabled = True
End Sub

Private Sub Check1_Click()
    If Check1 Then
        txtWasted.Enabled = True
        txtReturn.Enabled = True
        Label8.Enabled = True
        Label9.Enabled = True
    Else
        txtWasted.Enabled = False
        txtReturn.Enabled = False
        Label8.Enabled = False
        Label9.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        txtBackdate.Visible = True
    Else
        txtBackdate.Visible = False
    End If
End Sub

Private Sub cmdCompute_Click()
    On Error GoTo Ouch
    
    Dim Temp1 As Double
    
    If txtPresentReading = "" Then
        MsgBox "Enter a Pump Litre value"
        Exit Sub
    End If
    
    If txtLR = "" Then
        MsgBox "Please select a pump"
        Exit Sub
    End If
    
    If Check2.Value = vbChecked Then
        If txtBackdate = "" Then
            MsgBox "Please enter a value for the product!", vbExclamation
            txtBackdate.SetFocus
            Exit Sub
        End If
        Dim dTemp As Double
        dTemp = CDbl(txtBackdate)
        txtCost = txtBackdate
    End If
            
    dblPR = CDbl(txtPresentReading)
    dblWaste = CDbl(txtWasted)
    dblLR = CDbl(txtLR)
    If dblPR < dblLR Then
        'check if the pump has been reset - ask the user
        If MsgBox("Has the pump been RESET?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            'retrieve the reset value, do some maths
            Temp1 = CDbl(Text2) - dblLR
            dblTemp = dblPR + Temp1
        Else
            MsgBox "Present Reading can not be lesser than previous reading", vbCritical
            Me.txtPresentReading.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
    Else
        dblTemp = dblPR - dblLR
    End If
'    dblTemp = dblPR - dblLR
    dblSold = dblTemp - dblWaste
    
    dblCost = CDbl(txtCost)
    dblAmount = dblSold * dblCost
    
    txtSold = dblSold
    txtAmount = dblAmount
    cmdOK.Enabled = True
    cmdCompute.Enabled = False
    Exit Sub
    
Ouch:
    If Err.Number = 13 Then
        Exit Sub
    End If
    'MsgBox Err.Number & " :: " & Err.Description
End Sub

Private Sub cmdOK_Click()
    Dim sSQL As String

    mess = ""
    mess = mess & "Is the following information correct:" & vbCrLf & vbCrLf
    mess = mess & "Staff Name:" & Me.Text1 & vbCrLf & vbCrLf
    mess = mess & "Pump Number:" & Me.cboPump & vbCrLf & vbCrLf
    mess = mess & "Product Type:" & Me.txtPT & vbCrLf & vbCrLf
    mess = mess & "Present Reading:" & Me.txtPresentReading & vbCrLf & vbCrLf
    mess = mess & "Litres Sold: " & Me.txtSold & vbCrLf & vbCrLf
    mess = mess & "Litres Return to Tank: " & Me.txtWasted & vbCrLf & vbCrLf
    mess = mess & "Amount:" & Me.txtAmount & vbCrLf & vbCrLf
    If Val(txtWasted.Text) <> 0 Then
        mess = mess & "Reason for Return to Tank:" & Me.txtReturn & vbCrLf & vbCrLf
    Else
        mess = mess & "Not returned to Tank" & vbCrLf & vbCrLf
    End If
    If MsgBox(mess, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    cmdCompute.Enabled = False
    With frmSales
        .txtSold = Format(dblSold, "#0.00")
        .txtAmount = Format(dblAmount, "#0.00")
        .txtPumpID = cboPump.Text
        .Text4 = txtPT.Text
    End With
    
    'On Error GoTo Ouch
    'open database for transaction
    cn.BeginTrans
    
    'changes go here
    sSQL = "update tblPumpType set tblLastReading = " & CDbl(Me.txtPresentReading) & " where tblPumpID = " & cboPump.Text
    cn.Execute sSQL
    
    sSQL = "insert into tblPumpRecord(tblPumpID, tblDate, tblShift, tblStaffID, tblInitialLitres, tblFinalLitres, tblWasteLitres, tblLitresSold, tblUnitCost, tblTotalCost, tblReturn, tblStaffName) values ('" & cboPump.Text & "','" & frmSales.DTPicker1.Value & "','" & frmSales.cboShift.Text & "','" & frmSales.cboStaffID.Text & "'," & Me.txtLR & "," & Me.txtPresentReading & "," & Me.txtWasted & "," & Me.txtSold & "," & Me.txtCost & "," & Me.txtAmount & ",'" & Me.txtReturn & "','" & Me.Text1 & "')"
    cn.Execute sSQL
    
    
    'commit changes to database
    cn.CommitTrans
    MsgBox "Record Saved"
    
    Unload Me
    
    Exit Sub
    
Ouch:
    MsgBox Err.Number & Err.Description
    cn.RollbackTrans
    
    Exit Sub
    
dError:
    MsgBox "please enter a numeric value", vbCritical
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    sSQL = "select tblPumpID from tblPumpType"
    Set rsPumps = cn.Execute(sSQL)
    cboPump.Clear
    With rsPumps
        Do While Not .EOF
            cboPump.AddItem .Fields("tblPumpID")
            .MoveNext
        Loop
    End With
    
    frmSales.Enabled = False
    disableMainMenu False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSales.Enabled = True
    frmSales.Command2.Enabled = True
    disableMainMenu True
End Sub
