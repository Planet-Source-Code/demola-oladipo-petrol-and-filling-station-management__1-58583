VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmShiftDelivery1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift Delivery Report"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4350
   Begin RichTextLib.RichTextBox RichTextBox4 
      Height          =   735
      Left            =   8880
      TabIndex        =   6
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"Form2.frx":030A
   End
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   735
      Left            =   8880
      TabIndex        =   5
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"Form2.frx":0384
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   735
      Left            =   8760
      TabIndex        =   4
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"Form2.frx":03FE
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   9240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"Form2.frx":0478
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Report"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   69730305
      CurrentDate     =   37917
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   69730305
      CurrentDate     =   37917
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Shift"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmShiftDelivery1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSQL As String
Dim sSQL1 As String

Dim recCOunt As Integer

Dim rs As New Recordset
Dim rs1 As New Recordset

Private Sub Command1_Click()
    'On Error GoTo Ouch
    
    Dim Date1 As String
    Dim Date2 As String
    
    If Me.combo1.Text = "" Then
        MsgBox "Please select a Shift!", vbCritical
        Exit Sub
    End If
    
    
    
    Dim sShift As String
    Dim sShift1 As String
    

Select Case combo1.Text
    Case "Morning"
        sShift = "Morning"
        sSQL = "Select * from tblPumpRecord WHERE tblShift = 'Morning' and (tblDate>#" & DateAdd("d", 0, Format(DTPicker1.Value, "mm/dd/yyyy")) & "# And tblDate < #" & Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#)"
        Set rs = cn.Execute(sSQL)

    Case "Night"
        sShift = "Night"
        sSQL = "Select * from tblPumpRecord WHERE tblShift = 'Night' and (tblDate>#" & DateAdd("d", 0, Format(DTPicker1.Value, "mm/dd/yyyy")) & "# And tblDate < #" & Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#)"
        Set rs = cn.Execute(sSQL)
        
    Case "Both Shifts"
        sShift = "Morning"
        sSQL = "Select * from tblPumpRecord WHERE tblShift = 'Morning' and (tblDate>#" & DateAdd("d", 0, Format(DTPicker1.Value, "mm/dd/yyyy")) & "# And tblDate < #" & Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#)"
        Set rs = cn.Execute(sSQL)
        
        sShift = "Night"
        sSQL1 = "Select * from tblPumpRecord WHERE tblShift = 'Night' and (tblDate>#" & DateAdd("d", 0, Format(DTPicker1.Value, "mm/dd/yyyy")) & "# And tblDate < #" & Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#)"
        Set rs1 = cn.Execute(sSQL1)
        
'    Case "Night"
'        sShift = "Night"
'        sSQL = "SELECT tblSalesRecord.tblShiftPump, tblSalesRecord.tblShift, tblPumpRecord.tblInitialLitres, tblPumpRecord.tblFinalLitres, tblPumpRecord.tblWasteLitres, tblsalesRecord.tblshiftLitres, tblPumpRecord.tblUnitCost, tblPumpRecord.tblTotalCost, tblSalesRecord.tblDate AS tblSalesRecord_tblDate, tblSalesRecord.tblStaffID AS tblSalesRecord_tblStaffID, tblPumpRecord.tblStaffName, tblPumpRecord.tblReturn FROM tblStaff INNER JOIN (tblSalesRecord INNER JOIN tblPumpRecord ON [tblSalesRecord].[tblShiftPump]=[tblPumpRecord].[tblPumpID]) ON [tblStaff].[tblStaffID]=[tblSalesRecord].[tblStaffID] WHERE ((([tblSalesRecord].[tblShift])='" & sShift & "') And (([tblSalesRecord].[tblDate])>#" & _
'            Format(DTPicker1.Value, "mm/dd/yyyy") & "# And ([tblSalesRecord].[tblDate])<#" & _
'            Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#) And (([tblPumpRecord].[tblShift])='" & sShift & "') And (([tblPumpRecord].[tblDate])>#" & _
'            Format(DTPicker1.Value, "mm/dd/yyyy") & "# And ([tblPumpRecord].[tblDate])<#" & _
'            Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#));"
'        Set rs = cn.Execute(sSQL)
'
'    Case "Both Shifts"
'        sShift = "Morning"
'        sSQL = "SELECT tblSalesRecord.tblShiftPump, tblSalesRecord.tblShift, tblPumpRecord.tblInitialLitres, tblPumpRecord.tblFinalLitres, tblPumpRecord.tblWasteLitres, tblsalesRecord.tblshiftLitres, tblPumpRecord.tblUnitCost, tblPumpRecord.tblTotalCost, tblSalesRecord.tblDate AS tblSalesRecord_tblDate, tblSalesRecord.tblStaffID AS tblSalesRecord_tblStaffID, tblPumpRecord.tblStaffName, tblPumpRecord.tblReturn FROM tblStaff INNER JOIN (tblSalesRecord INNER JOIN tblPumpRecord ON [tblSalesRecord].[tblShiftPump]=[tblPumpRecord].[tblPumpID]) ON [tblStaff].[tblStaffID]=[tblSalesRecord].[tblStaffID] WHERE ((([tblSalesRecord].[tblShift])='" & sShift & "') And (([tblSalesRecord].[tblDate])>#" & _
'            Format(DTPicker1.Value, "mm/dd/yyyy") & "# And ([tblSalesRecord].[tblDate])<#" & _
'            Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#) And (([tblPumpRecord].[tblShift])='" & sShift & "') And (([tblPumpRecord].[tblDate])>#" & _
'            Format(DTPicker1.Value, "mm/dd/yyyy") & "# And ([tblPumpRecord].[tblDate])<#" & _
'            Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#));"
'        Set rs = cn.Execute(sSQL)
'
'
'        sShift = "Night"
'        sSQL1 = "SELECT tblSalesRecord.tblShiftPump, tblSalesRecord.tblShift, tblPumpRecord.tblInitialLitres, tblPumpRecord.tblFinalLitres, tblPumpRecord.tblWasteLitres, tblsalesRecord.tblshiftLitres, tblPumpRecord.tblUnitCost, tblPumpRecord.tblTotalCost, tblSalesRecord.tblDate AS tblSalesRecord_tblDate, tblSalesRecord.tblStaffID AS tblSalesRecord_tblStaffID, tblPumpRecord.tblStaffName, tblPumpRecord.tblReturn FROM tblStaff INNER JOIN (tblSalesRecord INNER JOIN tblPumpRecord ON [tblSalesRecord].[tblShiftPump]=[tblPumpRecord].[tblPumpID]) ON [tblStaff].[tblStaffID]=[tblSalesRecord].[tblStaffID] WHERE ((([tblSalesRecord].[tblShift])='" & sShift & "') And (([tblSalesRecord].[tblDate])>#" & _
'            Format(DTPicker1.Value, "mm/dd/yyyy") & "# And ([tblSalesRecord].[tblDate])<#" & _
'            Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#) And (([tblPumpRecord].[tblShift])='" & sShift & "') And (([tblPumpRecord].[tblDate])>#" & _
'            Format(DTPicker1.Value, "mm/dd/yyyy") & "# And ([tblPumpRecord].[tblDate])<#" & _
'            Format(DateAdd("d", 1, DTPicker2.Value), "mm/dd/yyyy") & "#));"
'        Set rs1 = cn.Execute(sSQL1)

End Select


If rs.BOF And rs.EOF Then
    If combo1.Text = "Both Shifts" Then
        If rs1.BOF And rs1.EOF Then
            MsgBox "No Record"
            Exit Sub
        End If
    Else
        MsgBox "No Record"
        Exit Sub
    End If
End If

createHTML
Load frmBrowser
With frmBrowser
    .brwWebBrowser.Navigate App.Path & "\temp.html"
    .Show
End With
Exit Sub

Ouch:
    MsgBox Err.Number & Err.Description
End Sub

Private Sub Form_Load()
    On Error Resume Next
    combo1.Clear
    combo1.AddItem "Morning"
    combo1.AddItem "Night"
    combo1.AddItem "Both Shifts"
    
    DTPicker1.Value = Now
    DTPicker2.Value = Now
End Sub

Sub createHTML()
Dim sDAte, sShift As String
Dim temp As Integer
Dim dblPMS As Double, dblAGO As Double, dblDPK As Double, dblAmount As Double, dblBULK As Double
'    Open "temp.html" For Output As #1
    On Error GoTo BugOut
    Kill App.Path & "\temp.html"
    RichTextBox1.LoadFile App.Path & "\res\shiftDe01.dll"
    RichTextBox2.LoadFile App.Path & "\res\shiftDe02.dll"
    RichTextBox3.LoadFile App.Path & "\res\shiftDe03.dll"
    sDAte = "Date:<b>" & Format(DTPicker1.Value, "dd-mmm-yy") & "</b>"
    sShift = "Shift: <b>" & combo1.Text & "</b>"
    
    RichTextBox4.Text = ""
    dblDPK = 0
    dblPMS = 0
    dblAGO = 0
    dblBULK = 0
    dblAmount = 0
    With rs
        temp = 0
        Do While Not .EOF
            DoEvents
            RichTextBox4.Text = RichTextBox4.Text & "<tr>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblStaffID") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblStaffName") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblPumpId") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblInitialLitres") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblFinalLitres") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblWasteLitres") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & Format(.Fields("tblLitresSold"), "#,##0.0") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblUnitCost") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblTotalCost") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblReturn") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "</tr>"
            
            dblAmount = dblAmount + CDbl(.Fields("tblTotalCost"))
            
            sSQL = "select tblPumpType from tblPumpType where tblPumpID = " & .Fields("tblPumpID")
            Set rsPumps = cn.Execute(sSQL)
            
            If rsPumps.BOF And rsPumps.EOF Then
                GoTo BugOut
            End If
            
            If rsPumps.Fields(0) = "PMS" Then
                dblPMS = dblPMS + CDbl(.Fields("tblLitresSold"))
            ElseIf rsPumps.Fields(0) = "AGO" Then
                dblAGO = dblAGO + CDbl(.Fields("tblLitresSold"))
            ElseIf rsPumps.Fields(0) = "DPK" Then
                dblDPK = dblDPK + CDbl(.Fields("tblLitresSold"))
            ElseIf rsPumps.Fields(0) = "BULK" Then
                dblBULK = dblBULK + CDbl(.Fields("tblLitresSold"))
            Else
                'MsgBox " ", vbCritical
            End If
            .MoveNext
            temp = temp + 1
        Loop
    End With
    With rs1
        Do While Not .EOF
            RichTextBox4.Text = RichTextBox4.Text & "<tr>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblStaffID") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblStaffName") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblPumpId") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblInitialLitres") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblFinalLitres") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblWasteLitres") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & Format(.Fields("tblLitresSold"), "#,##0.0") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblUnitCost") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblTotalCost") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields("tblReturn") & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "</tr>"
'            RichTextBox4.Text = RichTextBox4.Text & "<tr>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(9) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(10) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(0) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(2) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(3) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(4) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & Format(.Fields(5), "#,##0.0") & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(6) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(7) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(11) & "</td>"
'            RichTextBox4.Text = RichTextBox4.Text & "</tr>"
            
            dblAmount = dblAmount + CDbl(.Fields("tblTotalCost"))
            
            sSQL = "select tblPumpType from tblPumpType where tblPumpID = " & .Fields(0)
            Set rsPumps = cn.Execute(sSQL)
            
            If rsPumps.BOF And rsPumps.EOF Then
                GoTo BugOut
            End If
            
            If rsPumps.Fields(0) = "PMS" Then
                dblPMS = dblPMS + CDbl(.Fields("tblLitresSold"))
            ElseIf rsPumps.Fields(0) = "AGO" Then
                dblAGO = dblAGO + CDbl(.Fields("tblLitresSold"))
            ElseIf rsPumps.Fields(0) = "DPK" Then
                dblDPK = dblDPK + CDbl(.Fields("tblLitresSold"))
            ElseIf rsPumps.Fields(0) = "BULK" Then
                dblBULK = dblBULK + CDbl(.Fields("tblLitresSold"))
            Else
                'MsgBox " ", vbCritical
            End If
            .MoveNext
            temp = temp + 1
        Loop
    End With
HTML:
    RichTextBox1.Text = RichTextBox1.Text & Format(DTPicker1.Value, "dd-mmm-yy") & " to " & Format(DTPicker2.Value, "dd-mmm-yy") & " and " & sShift & RichTextBox2.Text & RichTextBox4.Text & RichTextBox3.Text
    RichTextBox1.Text = RichTextBox1.Text & "Total Amount: <b>N" & Format(dblAmount, "#,##0.00") & "</b><br>" & _
        "Total PMS Volumes sold: <b>" & Format(dblPMS, "#,##0.0") & "</b> &nbsp; &nbsp; &nbsp; " & _
        "Total AGO Volumes sold: <b>" & Format(dblAGO, "#,##0.0") & "</b> &nbsp; &nbsp; &nbsp;" & _
        "Total DPK Volumes sold: <b>" & Format(dblDPK, "#,##0.0") & "</b> &nbsp; &nbsp; &nbsp;" & _
        "Total BULK Oil Volumes sold: <b>" & Format(dblBULK, "#,##0.0") & "</b> &nbsp; &nbsp; &nbsp; " & "</font></td></tr></table></body></html>"
    
    RichTextBox1.SaveFile App.Path & "\temp.html", rtfText
    MsgBox temp & " records found in the database!"
    Exit Sub
    
BugOut:
    If Err.Number = 53 Then
        Resume Next
    ElseIf Err.Number = 3704 Then
        Resume HTML
    Else
        MsgBox Err.Number & Err.Description
    End If
    '    Resume HTML
End Sub
