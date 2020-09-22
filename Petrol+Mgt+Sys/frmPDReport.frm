VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPDReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Pumps Daily Report"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "frmPDReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3960
   Begin RichTextLib.RichTextBox RichTextBox4 
      Height          =   735
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmPDReport.frx":030A
   End
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   735
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmPDReport.frx":0395
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmPDReport.frx":0420
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmPDReport.frx":04AB
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Report"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65273857
      CurrentDate     =   37919
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Shift"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPDReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim sSQL As String

Private Sub Command1_Click()
Dim Day1 As String, Day2 As String

If Combo1.Text = "" Then
    MsgBox "Please select a Shift", vbCritical
    Exit Sub
End If
Day1 = Format(DateAdd("d", -1, DTPicker1.Value), "mm/dd/yyyy")
Day2 = Format(DateAdd("d", 1, DTPicker1.Value), "mm/dd/yyyy")
Daynow = Format(DTPicker1.Value, "mm/dd/yyyy")

sSQL = "SELECT tblPumpRecord.tblPumpID, tblPumpType.tblPumpType, tblPumpRecord.tblFinalLitres, " & _
"tblPumpRecord.tblWasteLitres, tblStaff.tblSurname, tblStaff.tblFirstName, tblPumpRecord.tblShift, " & _
"tblPumpRecord.tblDate, tblPumpRecord.tblInitialLitres FROM (tblPumpType RIGHT JOIN tblPumpRecord ON " & _
"tblPumpType.tblPumpID = tblPumpRecord.tblPumpID) " & _
"LEFT JOIN tblStaff ON tblPumpRecord.tblStaffId = tblStaff.tblStaffID WHERE " & _
"(((tblPumpRecord.tblShift)='" & Combo1.Text & "') AND ((tblPumpRecord.tblDate) > #" & Daynow & "#) " & _
"and ((tblPumpRecord.tblDate) < #" & Day2 & "#))"

Set rs = cn.Execute(sSQL)
If rs.BOF And rs.EOF Then
    MsgBox "No Record"
    Exit Sub
End If

createHTML
Load frmBrowser
With frmBrowser
    .brwWebBrowser.Navigate App.Path & "\temp1.html"
    .Show
End With
Exit Sub

Ouch:
    MsgBox Err.Number & Err.Description
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Combo1.Clear
    Combo1.AddItem "Morning"
    Combo1.AddItem "Night"
End Sub

Sub createHTML()
    Dim sDAte, sShift As String
    On Error Resume Next
    Kill App.Path & "\temp1.html"
    RichTextBox1.LoadFile App.Path & "\res\pd01.dll"
    RichTextBox2.LoadFile App.Path & "\res\pd02.dll"
    RichTextBox3.LoadFile App.Path & "\res\pd03.dll"
    sDAte = "Date:<b>" & Format(DTPicker1.Value, "dd-mmm-yy") & "</b>"
    sShift = "Shift: <b>" & Combo1.Text & "</b>"
    
    RichTextBox4.Text = ""
    With rs
        Do While Not .EOF
            RichTextBox4.Text = RichTextBox4.Text & "<tr>"
            'RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(7) & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(0) & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(1) & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(8) & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(2) & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(3) & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "<td>" & .Fields(4) & " " & .Fields(5) & "</td>"
            RichTextBox4.Text = RichTextBox4.Text & "</tr>"
            .MoveNext
        Loop
    End With
HTML:
    RichTextBox1.Text = RichTextBox1.Text & sDAte & " and " & sShift & RichTextBox2.Text & RichTextBox4.Text & RichTextBox3.Text
    'RichTextBox1.Text = RichTextBox1.Text & "Total Amount: <b>N" & Format(dblAmount, "#,##0.00") & "</b><br>" & "Total PMS Volumes sold: <b>" & Format(dblPMS, "#,##0") & "</b> &nbsp; &nbsp; &nbsp; " & "Total AGO Volumes sold: <b>" & Format(dblAGO, "#,##0") & "</b> &nbsp; &nbsp; &nbsp;" & "Total DPK Volumes sold: <b>" & Format(dblDPK, "#,##0") & "</b> &nbsp; &nbsp; &nbsp; " & "</font></td></tr></table></body></html>"
    RichTextBox1.SaveFile App.Path & "\temp1.html", rtfText
    
    Exit Sub
    
BugOut:
    Resume HTML
End Sub

