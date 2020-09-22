VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salary Report"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "frmSalary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4035
   Begin VB.CommandButton Command1 
      Caption         =   "View Report"
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   37913
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   37913
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "End Date"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Date"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1, rs2 As New Recordset
Dim sSQL As String
Dim rsTemp As New Recordset

Private Sub Command1_Click()
    'check for data relevance
    If DTPicker1.Value > DTPicker2 Then
        MsgBox "Starting Date can not be less than End Date", vbCritical
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    'empty the recycle bin
    sSQL = "delete from tbldump"
    cn1.Execute sSQL
    
    Dim intP, intA, intOff, intTargetCount As Integer
    Dim intOffA, intSalesTarget, intSalesPresent As Integer
    Dim dblTarget, dblBonus, dblSalary As Double
    Dim dblSalesBonus1, dblSalesBonus2, dblSalesBonus3 As Double
    Dim pmsSales As Double
    
    intOffA = CInt(GetStringSetting(strProdName, "Sales Records", "OffPlusAbsent", "4"))
    intSalesTarget = CInt(GetStringSetting(strProdName, "Sales Records", "Sales Target", "20"))
    intSalesPresent = CInt(GetStringSetting(strProdName, "Sales Records", "Sales Present", "20"))
    dblSalesBonus1 = CDbl(GetStringSetting(strProdName, "Sales Records", "Bonus1", "5000"))
    dblSalesBonus2 = CDbl(GetStringSetting(strProdName, "Sales Records", "Bonus2", "1400"))
    dblSalesBonus3 = CDbl(GetStringSetting(strProdName, "Sales Records", "Bonus3", "0"))
    
    'cycle through all records checking within current date range
    For i = 0 To List1.ListCount - 1
        sSQL = "select * from tblSalesrecord where tblStaffID = '" & List1.List(i) & "' and tblDate Between #" & Format(DTPicker1.Value, "dd-mmm-yy") & "# And #" & Format(DateAdd("d", 1, DTPicker2.Value), "dd-mmm-yy") & "#"
        Set rs1 = cn.Execute(sSQL)
        If rs1.EOF And rs1.BOF Then
            'no record for that staff for that period
            GoTo Label1
        End If
        With rs1
            intP = 0
            intA = 0
            intOff = 0
            intTargetCount = 0
            dblTarget = 0
            dblBonus = 0
            dblSalary = 0
            pmsSales = 0
            
            Do While Not .EOF
                'MsgBox .Fields("tblDate") & .Fields("tblStaffID") & .Fields("tblAmount")
                If .Fields("tblAttendance") = "Present" Then
                    intP = intP + 1
                ElseIf .Fields("tblAttendance") = "Absent" Then
                    intA = intA + 1
                ElseIf .Fields("tblAttendance") = "Off" Then
                    intOff = intOff + 1
                ElseIf .Fields("tblAttendance") = "Off with Permission" Then
                    'do nothing
                    'intOff = intOff + 1
                'Else
                    
                End If
                If CInt(.Fields("tblTarget")) <> 0 Then
                    intTargetCount = intTargetCount + 1
                End If
                dblSalary = dblSalary + .Fields("tblTarget") + .Fields("tblIncentive")
                
                On Error Resume Next
                'check for sales type
                sSQL = "select tblPumpType from tblPumpType where tblPumpID = " & CInt(.Fields("tblShiftPump"))
                Set rsTemp = cn.Execute(sSQL)
                Select Case rsTemp.Fields(0).Value
                    Case "PMS"
                        pmsSales = pmsSales + CDbl(.Fields("tblShiftLitres"))
                End Select
            
                .MoveNext
            Loop
            'check criteria for sales records
            
            If ((intOff + intA <= intOffA) And (intTargetCount >= intSalesTarget) And (intP >= intSalesPresent)) Then
                dblBonus = dblSalesBonus1
            ElseIf ((intOff + intA <= 4) And (intTargetCount <= 20) And (intP >= 20)) Then
                dblBonus = dblSalesBonus2
            ElseIf ((intOff + intA <= 4) And (intTargetCount <= 20) And (intP <= 20)) Then
                dblBonus = dblSalesBonus3
            Else
                'this is an error situation
            End If
            
            'add bonus to salary
            dblSalary = dblSalary + dblBonus
            
            
            
            On Error Resume Next
            
            'retrieve staff name
            Dim strSName As String
            sSQL = "select tblSurname, tblFirstName from tblStaff where tblStaffId = '" & List1.List(i) & "'"
            Set rsStaff = cn.Execute(sSQL)
            strSName = rsStaff.Fields(0) & " " & rsStaff.Fields(1)
            rsStaff.Close
            
            'retrieve staff name from suspended
            'Dim strSName As String
            sSQL = "select tblSurname, tblFirstName from tblSuspend where tblStaffId = '" & List1.List(i) & "'"
            Set rsStaff = cn.Execute(sSQL)
            strSName = rsStaff.Fields(0) & " " & rsStaff.Fields(1)
            'rsStaff.Close
            
            'retrieve staff name from retrenched
            'Dim strSName As String
            sSQL = "select tblSurname, tblFirstName from tblRetrench where tblStaffId = '" & List1.List(i) & "'"
            Set rsStaff = cn.Execute(sSQL)
            strSName = rsStaff.Fields(0) & " " & rsStaff.Fields(1)
            'rsStaff.Close
            
            'On Error GoTo 0
            
            'add to refuse bin
            
            sSQL = "insert into tblDump(tblStaffID, tblStaffName, tblPResent, tblAbsent, tblOff, tblTarget, tblBonus, tblSalary, tblSales) values ('" & List1.List(i) & "','" & strSName & "'," & intP & "," & intA & "," & intOff & "," & intTargetCount & "," & dblBonus & "," & dblSalary & "," & pmsSales & ")"
            cn1.Execute sSQL
            strSName = ""
        End With
Label1:
    DoEvents
    Next
    
    With dEnvDump.rscmdSalary
        If .State Then
            .Close
            Unload dRptSalary
        End If
    End With
    Load dRptSalary
    With dRptSalary
        .Caption = "Staff Salary Report for " & Format(DTPicker1.Value, "mmmm-dd-yyyy") & " and " & Format(DTPicker2.Value, "mmmm-dd-yyyy")
        .Show
    End With
    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    'retrieve all current staff id numbers into the listbox
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    sSQL = "select tblStaffID from tblStaff"
    Set rs1 = cn.Execute(sSQL)
    With rs1
        Do While Not .EOF
            List1.AddItem .Fields("tblStaffID")
            .MoveNext
        Loop
    End With
    
    'retrieve all suspended staff id numbers into the listbox
    
    sSQL = "select tblStaffID from tblSuspend"
    Set rs1 = cn.Execute(sSQL)
    With rs1
        Do While Not .EOF
            List1.AddItem .Fields("tblStaffID")
            .MoveNext
        Loop
    End With
    
    'retrieve all retrenched staff id numbers into the listbox
    
    sSQL = "select tblStaffID from tblRetrench"
    Set rs1 = cn.Execute(sSQL)
    With rs1
        Do While Not .EOF
            List1.AddItem .Fields("tblStaffID")
            .MoveNext
        Loop
    End With
    
    
End Sub
