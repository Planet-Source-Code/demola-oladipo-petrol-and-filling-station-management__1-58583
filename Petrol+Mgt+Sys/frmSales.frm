VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdUB 
      Caption         =   "UnderBonnet"
      Height          =   735
      Left            =   5640
      TabIndex        =   36
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5760
      TabIndex        =   35
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox cboStaffID 
      Height          =   2595
      Left            =   5760
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPumpID 
      Height          =   615
      Left            =   5760
      TabIndex        =   33
      Text            =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Record"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5640
      TabIndex        =   32
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sales Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtSold 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate"
         Height          =   735
         Left            =   3960
         TabIndex        =   28
         Top             =   360
         Width           =   1335
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
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Litres Sold"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Bonus Computation"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   5415
      Begin VB.TextBox txtBonus 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """N""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtTIncentive 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """N""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtAIncentive 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """N""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Compute"
         Height          =   975
         Left            =   3960
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Total Bonus"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Target Incentive"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Attendance Incentive"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shifts"
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ComboBox cboShift 
         Height          =   315
         ItemData        =   "frmSales.frx":030A
         Left            =   120
         List            =   "frmSales.frx":0314
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "cboShift"
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Date and Time"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   2775
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   69795841
         CurrentDate     =   37912
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Present/Absent"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   2535
      Begin VB.OptionButton optOff 
         Caption         =   "Off"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optPM 
         Caption         =   "Absent with Permission"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optAbsent 
         Caption         =   "Absent"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optPresent 
         Caption         =   "Present"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Staff Identification"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   2895
      End
      Begin VB.ComboBox cboStaffID1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Sex"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Firstname"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Surname"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Staff Name"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As New Recordset

Private Sub cboShift_Click()
    If Not staffCleared Then
        With Frame6
            .Enabled = False
            .Visible = False
            Command1.Enabled = False
        End With
        Exit Sub
    End If
    
    'sSQL = "select * from
    If cboShift.Text <> "" Then
        With Frame6
            .Enabled = True
            .Visible = True
            Command1.Enabled = True
        End With
        cmdUB.Enabled = True
        cmdUB.Visible = True
    Else
        With Frame6
            .Enabled = False
            .Visible = False
            Command1.Enabled = False
        End With
    End If
End Sub


Private Sub cboStaffID1_Change()
    cboStaffID1_Click
End Sub

Private Sub cboStaffID1_Click()
    On Error Resume Next
    Dim sSQL As String
    cboStaffID.ListIndex = cboStaffID1.ListIndex
    sSQL = "select * from tblStaff where tblStaffId = '" & cboStaffID.Text & "'"
    Set rs = cn.Execute(sSQL)
    Text1 = rs.Fields("tblSurname")
    Text2 = rs.Fields("tblfirstname")
    Text3 = rs.Fields("tblSex")
    enableFrames True
    Frame6.Visible = False
    Frame5.Visible = False
    Frame3.Visible = False
    Frame4.Enabled = True
    cmdSave.Enabled = False
    
    Me.optAbsent = False
    Me.optOff = False
    Me.optPM = False
    Me.optPresent = False
    Me.txtAIncentive = 0
    Me.txtAmount = 0
    Me.txtBonus = 0
    Me.txtPumpID = 0
    Me.txtSold = 0
    Me.txtTIncentive = 0
End Sub


Private Sub cmdSave_Click()
    On Error GoTo Ouch
    Dim sSQL, strAttend, sShift As String
    
    If Not staffCleared Then
        Exit Sub
    End If
    If Me.optAbsent Then
        strAttend = "Absent"
    ElseIf Me.optOff Then
        strAttend = "Off"
    ElseIf Me.optPM Then
        strAttend = "Off with Permission"
    ElseIf Me.optPresent Then
        strAttend = "Present"
    Else
        MsgBox "Please select an Attendance Value", vbCritical
        Exit Sub
    End If
    
    If optPresent Then
        If cboShift.Text = "" Then
            MsgBox "Please select a Shift Period", vbCritical
            Exit Sub
        Else
            sShift = cboShift.Text
        End If
    Else
        sShift = ""
    End If
    
   
    If Me.txtAIncentive = "" Then Me.txtAIncentive = 0
    If Me.txtTIncentive = "" Then Me.txtTIncentive = 0
    If Me.txtPumpID = "" Then Me.txtPumpID = 0
    If Me.txtSold = "" Then txtSold = 0
    If Me.txtAmount = "" Then txtAmount = 0
    
    '<--
    sSQL = "select tblStaffID, tblShift, tblDate from tblSalesRecord where tblStaffID = '" & cboStaffID.Text & _
    "' and tblDAte= #" & DTPicker1.Value & "#" & _
    " and tblShift = '" & sShift & "'"
    
    Set rsTemp = cn.Execute(sSQL)
    If Not (rsTemp.EOF And rsTemp.BOF) Then
        strAttend = ""
        txtAIncentive = 0
    End If
    '<--
    
    '
    'another debug statement placed here
    '
    'if shift is night, check for morning
    sSQL = "select tblStaffID, tblShift, tblDate from tblSalesRecord where tblStaffID = '" & cboStaffID.Text & _
    "' and tblDAte= #" & DTPicker1.Value & "#" & _
    " and tblShift = 'Morning'"
    Set rsTemp = cn.Execute(sSQL)
    If rsTemp.EOF And rsTemp.BOF Then
        'user has not done morning
        'this stmt does nothing
        strAttend = strAttend
    Else
        'user has done morning
        strAttend = ""
    End If
    'end debug
    
    sSQL = "insert into tblSalesRecord(tblDate, tblStaffID, tblAttendance, tblShift, tblShiftPump, tblShiftLitres, " _
        & "tblTarget, tblIncentive, tblAmount) values ('" _
        & DTPicker1.Value & "','" & Me.cboStaffID & "','" & strAttend & "','" & sShift & "'," & Me.txtPumpID & _
         "," & Me.txtSold & "," & Me.txtTIncentive & "," & Me.txtAIncentive & "," & Me.txtAmount & ")"

    mess = "Is the following information correct:" & vbCrLf
    mess = mess & "- Staff Information ------------------------" & vbCrLf
    mess = mess & "Name      : " & Me.Text1 & " " & Me.Text2 & vbCrLf
    mess = mess & "Date      : " & Format(DTPicker1.Value, "dd - mmm - yyyy") & vbCrLf
    mess = mess & "Shift     : " & sShift & vbCrLf
    If optPresent Then
        mess = mess & "Attendance: Present" & vbCrLf
    ElseIf optAbsent Then
        mess = mess & "Attendance: Absent" & vbCrLf
    ElseIf optOff Then
        mess = mess & "Attendance: Off" & vbCrLf
    ElseIf optPM Then
        mess = mess & "Attendance: Absent with Permission" & vbCrLf
    Else
        mess = mess & ""
    End If
    mess = mess & "- Sales Information -------------------------" & vbCrLf
    mess = mess & "Litres    : " & Me.txtSold & vbCrLf
    mess = mess & "Amount    : " & Me.txtAmount & vbCrLf
    mess = mess & "- Incentive Information ---------------------" & vbCrLf
    mess = mess & "Attendance: " & Me.txtAIncentive & vbCrLf
    mess = mess & "Target    : " & Me.txtTIncentive & vbCrLf
    mess = mess & "Bonus     : " & Me.txtBonus
    
    mess1 = MsgBox(mess, vbYesNo)
    If mess1 = vbYes Then
        cn.Execute sSQL
    Else
        Exit Sub
    End If
    
    
    MsgBox "Record Saved", vbInformation
    cmdSave.Enabled = False
    Frame3.Enabled = False
    optPresent = False
    optOff = False
    optPM = False
    optAbsent = False
    Command2.Enabled = False
    
    Exit Sub
Ouch:
    If Err.Number = -2147467259 Then
        MsgBox "You are trying to create duplicate entries into the system. The action has been cancelled", vbCritical
        Exit Sub
    Else
        MsgBox Err.Number & Err.Description
    End If
End Sub

Private Sub cmdUB_Click()
    Dim newDate As String
    Dim oldDate As String
    
    newDate = Format(DTPicker1.Value, "mmm-dd-yyyy")
    If MsgBox("Confirm: '" & Me.cboStaffID1.Text & "' as Underbonnet", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    'check for staff previous underbonnet
    
    sSQL = "select * from tblSalesRecord where tblStaffID = '" & cboStaffID1.Text & "'  and tblAttendance = 'Present' and tblShiftPump = 0 order by tblDate desc"
    Set rsTemp = cn.Execute(sSQL)
    If rsTemp.BOF And rsTemp.EOF Then
        GoTo lblOut
    Else
        oldDate = Format(rs.Fields("tblDate"), "mmm-dd-yyyy")
        If oldDate <> newDate Then
            GoTo lblOut
        Else
            MsgBox "Sorry... No cant do"
            Exit Sub
        End If
    End If

lblOut:
    sSQL = "select * from tblSalesRecord where tblStaffID = '" & cboStaffID1.Text & "'  and tblAttendance = 'Present' and tblShiftPump <> 0 order by tblDate desc"
    Set rsTemp = cn.Execute(sSQL)
    If rsTemp.BOF And rsTemp.EOF Then
        GoTo lblOut1
    Else
        oldDate = Format(rs.Fields("tblDate"), "mmm-dd-yyyy")
        If oldDate <> newDate Then
            GoTo lblOut1
        Else
            MsgBox "Sorry... No cant do"
            Exit Sub
        End If
    End If
    
lblOut1:
    Frame5.Visible = True
    
    txtAIncentive = "100"
    txtBonus = "100"
    cmdSave_Click
    
    cmdUB.Enabled = False
    Command1.Enabled = False
End Sub

Private Sub Command1_Click()
mess = "Are you sure you want to calculate litres sold for staff: " & Me.cboStaffID1.Text & "?"
If MsgBox(mess, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
    Exit Sub
End If
Frame3.Enabled = False
Frame2.Enabled = False
Frame4.Enabled = False
Load frmPumpLitre
frmPumpLitre.Text1.Text = cboStaffID1.Text
frmPumpLitre.Show
Frame5.Visible = True
Frame5.Enabled = True
Command1.Enabled = False
cmdUB.Enabled = False
End Sub

Public Sub Command2_Click()
    Dim dblBonus, dblBonus1 As Double
    If cboShift.Text = "" Then
        If optPresent Then
            MsgBox "Please select a Shift!", vbExclamation
            Exit Sub
        End If
    End If
    Dim ii As String
    If Text4 = "PMS" Then
        ii = "0"
    ElseIf Text4 = "AGO" Then
        ii = "1"
    ElseIf Text4 = "BULK" Then
        ii = "3"
    Else
        ii = "2"
    End If
    'check for female staff
    If LCase$(Text3) = "female" Then
        'check for attendance irrespective of shift
        If optAbsent Then
            dblBonus1 = CDbl(GetStringSetting(strProdName, "Options-Female", "Absent", "-50"))
        ElseIf optPM Then
            dblBonus1 = 0
        ElseIf optOff Then
            dblBonus1 = 0
        ElseIf optPresent Then
            dblBonus1 = CDbl(GetStringSetting(strProdName, "Options-Female", "Present", "100"))
            'only present staff can make sales
            'check for morning shift
            If cboShift.Text = "Morning" Then
                'if female staff in morning shift meets or exceeds target
                If CDbl(txtSold) >= CDbl(GetStringSetting(strProdName, "Options-Female", "Morning Target" & ii, "3300")) Then
                    dblBonus = CDbl(GetStringSetting(strProdName, "Options-Female", "Morning Incentive" & ii, "100"))
                Else
                    dblBonus = 0
                End If
            'else it is a night shift
            ElseIf cboShift.Text = "Night" Then
                'if female staff in night shift meets or exceeds target
                If CDbl(txtSold) >= CDbl(GetStringSetting(strProdName, "Options-Female", "Night Target" & ii, "4000")) Then
                    dblBonus = CDbl(GetStringSetting(strProdName, "Options-Female", "Night Incentive" & ii, "100"))
                End If
            End If
        End If
        txtTIncentive = dblBonus
        txtAIncentive = dblBonus1
        
        txtBonus = dblBonus + dblBonus1
    ElseIf LCase$(Text3) = "male" Then
        'check for attendance irrespective of shift
        If optAbsent Then
            dblBonus1 = CDbl(GetStringSetting(strProdName, "Options-Male", "Absent", "-50"))
        ElseIf optPM Then
            dblBonus1 = 0
        ElseIf optOff Then
            dblBonus1 = 0
        ElseIf optPresent Then
            dblBonus1 = CDbl(GetStringSetting(strProdName, "Options-Male", "Present", "100"))
            'only present staff can make sales
            'check for morning shift
            If cboShift.Text = "Morning" Then
                'if male staff in morning shift meets or exceeds target
                If CDbl(txtSold) >= CDbl(GetStringSetting(strProdName, "Options-Male", "Morning Target" & ii, "3000")) Then
                    dblBonus = CDbl(GetStringSetting(strProdName, "Options-Male", "Morning Incentive" & ii, "100"))
                Else
                    dblBonus = 0
                End If
            'else it is a night shift
            ElseIf cboShift.Text = "Night" Then
                'if male staff in night shift meets or exceeds target
                If CDbl(txtSold) >= CDbl(GetStringSetting(strProdName, "Options-Male", "Night Target" & ii, "4000")) Then
                    dblBonus = CDbl(GetStringSetting(strProdName, "Options-Male", "Night Incentive" & ii, "100"))
                End If
            End If
        End If
        
        If Text4 <> "PMS" Then
            dblBonus = 0
        End If
        
        txtTIncentive = dblBonus
        txtAIncentive = dblBonus1
        
        txtBonus = dblBonus + dblBonus1
    End If
    cmdSave.Enabled = True
    Command2.Enabled = False
End Sub


Private Sub Form_Load()
    Dim sSQL As String
    
    cboStaffID.Clear
    cboStaffID1.Clear
    
    'customer care
    sSQL = "select tblStaffID, tblSurname, tblFirstname from tblStaff"
    Set rsStaff = cn.Execute(sSQL)
    With rsStaff
        Do While Not .EOF
            cboStaffID.AddItem .Fields("tblStaffId")
            cboStaffID1.AddItem .Fields(1) & " " & .Fields(2)
            .MoveNext
        Loop
    End With
    
    'senior staff
    sSQL = "select tblStaffID, tblSurname, tblFirstname from tblSeniorStaff"
    Set rsStaff = cn.Execute(sSQL)
    With rsStaff
        Do While Not .EOF
            cboStaffID.AddItem .Fields("tblStaffId")
            cboStaffID1.AddItem .Fields(1) & " " & .Fields(2)
            .MoveNext
        Loop
    End With
    
    'security staff
    sSQL = "select tblStaffID, tblSurname, tblFirstname from tblSecurity"
    Set rsStaff = cn.Execute(sSQL)
    With rsStaff
        Do While Not .EOF
            cboStaffID.AddItem .Fields("tblStaffId")
            cboStaffID1.AddItem .Fields(1) & " " & .Fields(2)
            .MoveNext
        Loop
    End With
    
    DTPicker1.Value = Now
    enableFrames False
End Sub

Private Sub enableFrames(eMode As Boolean)
    Frame2.Enabled = eMode
    Frame3.Enabled = eMode
    Frame4.Enabled = eMode
    Frame5.Enabled = eMode
    Frame6.Enabled = eMode
    
    Frame2.Visible = eMode
    Frame3.Visible = eMode
    Frame4.Visible = eMode
    Frame5.Visible = eMode
    Frame6.Visible = eMode
End Sub

Private Sub optAbsent_Click()
    With Frame6
        .Enabled = False
        .Visible = False
    End With
    With Frame3
        .Enabled = False
        .Visible = False
    End With
    cboShift.Text = ""
    cboShift.Locked = True
    cmdSave.Enabled = True
    
    cmdUB.Enabled = False
    cmdUB.Visible = False
    
    Me.Command2_Click
End Sub

Private Sub optOff_Click()
    With Frame6
        .Enabled = False
        .Visible = False
    End With
    With Frame3
        .Enabled = False
        .Visible = False
    End With
    cboShift.Text = ""
    cboShift.Locked = True
    cmdSave.Enabled = True
    
    cmdUB.Enabled = False
    cmdUB.Visible = False
    
    Me.Command2_Click
End Sub

Private Sub optPM_Click()
    With Frame6
        .Enabled = False
        .Visible = False
    End With
    With Frame3
        .Enabled = False
        .Visible = False
    End With
    cboShift.Text = ""
    cboShift.Locked = True
    cmdSave.Enabled = True
    
    cmdUB.Enabled = False
    cmdUB.Visible = False
    
    Me.Command2_Click
End Sub

Private Sub optPresent_Click()
    With Frame3
        .Enabled = True
        .Visible = True
    End With
    cboShift.Text = ""
    cboShift.Locked = False
    cmdSave.Enabled = False
    
    cmdUB.Enabled = False
    cmdUB.Visible = False
    
    'Me.Command2_Click
End Sub

Function staffCleared() As Boolean
    Dim oldDate As String
    Dim newDate As String

    staffCleared = False
    '
    ' -- check for date entry by format dates to be alike
    newDate = Format(DTPicker1.Value, "dd-mmm-yyyy")
    
    '
    ' check for staff as present
    sSQL = "select * from tblSalesRecord where tblStaffID = '" & cboStaffID.Text & "' and tblAttendance = 'Present' and tblShiftPump <> 0 order by tblDate desc"
    Set rsTemp = cn.Execute(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
        GoTo lblOkay1
    Else
        oldDate = Format(rsTemp.Fields("tblDate"), "dd-mmm-yyyy")
        If oldDate <> newDate Then
            GoTo lblOkay1
        Else
            If MsgBox("'" & cboStaffID1.Text & "' has been registered today as Present and made sales. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Function
            Else
                GoTo lblOkay1
            End If
        End If
    End If
lblOkay1:
    
    '
    ' do a secondary check
    sSQL = "select * from tblSalesRecord where tblStaffID = '" & cboStaffID.Text & "' and tblAttendance = 'Present' and tblShiftPump = 0 order by tblDate desc"
    Set rsTemp = cn.Execute(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
        GoTo lblOkay
    Else
        oldDate = Format(rsTemp.Fields("tblDate"), "dd-mmm-yyyy")
        If oldDate <> newDate Then
            GoTo lblOkay
        Else
            MsgBox "'" & cboStaffID1.Text & "' has been registered today as Present and Underbonnet.", vbCritical
            Exit Function
        End If
    End If
lblOkay:

    '
    ' check for staff as present
    sSQL = "select * from tblSalesRecord where tblStaffID = '" & cboStaffID.Text & "' and tblAttendance = 'Off' order by tblDate desc"
    Set rsTemp = cn.Execute(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
        GoTo lblOkay2
    Else
        oldDate = Format(rsTemp.Fields("tblDate"), "dd-mmm-yyyy")
        If oldDate <> newDate Then
            GoTo lblOkay2
        Else
            If MsgBox("'" & cboStaffID1.Text & "' has been registered today as Off. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Function
            Else
                GoTo lblOkay2
            End If
        End If
    End If
lblOkay2:
    
    '
    ' do a secondary check
    sSQL = "select * from tblSalesRecord where tblStaffID = '" & cboStaffID.Text & "' and tblAttendance = 'Absent' order by tblDate desc"
    Set rsTemp = cn.Execute(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
        GoTo lblOkay3
    Else
        oldDate = Format(rsTemp.Fields("tblDate"), "dd-mmm-yyyy")
        If oldDate <> newDate Then
            GoTo lblOkay3
        Else
            If MsgBox("'" & cboStaffID1.Text & "' has been registered today as Absent. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Function
            Else
                GoTo lblOkay3
            End If
        End If
    End If
lblOkay3:

    '
    ' do a secondary check
    sSQL = "select * from tblSalesRecord where tblStaffID = '" & cboStaffID.Text & "' and tblAttendance = 'Off with Permission' order by tblDate desc"
    Set rsTemp = cn.Execute(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
        GoTo lblOkay4
    Else
        oldDate = Format(rsTemp.Fields("tblDate"), "dd-mmm-yyyy")
        If oldDate <> newDate Then
            GoTo lblOkay4
        Else
            If MsgBox("'" & cboStaffID1.Text & "' has been registered today as Absent with Permission. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Function
            Else
                GoTo lblOkay4
            End If
        End If
    End If
lblOkay4:

staffCleared = True
End Function
