VERSION 5.00
Begin VB.Form frmDelete1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Staff Record..."
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmDelete1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Width           =   3375
   End
   Begin VB.ListBox combo1 
      Height          =   2205
      Left            =   4440
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Pump:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Attendance:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Staff Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmDelete1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
    sSQL = "select tblAttendance, tblDate,tblShiftPump from tblSalesRecord where tblStaffID = '" & combo1.List(Combo2.ListIndex) & "' order by tblDate desc"
    Set rs = cn.Execute(sSQL)
    If rs.BOF And rs.EOF Then
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Command1.Enabled = False
        Exit Sub
    End If
    If rs.Fields("tblAttendance") <> "Present" Then
        Text1.Text = rs.Fields("tblAttendance")
        Text2.Text = rs.Fields("tblDate")
        Text3.Text = rs.Fields("tblShiftPump")
        
        Command1.Enabled = True
    ElseIf rs.Fields("tblAttendance") = "Present" And rs.Fields("tblShiftPump") = 0 Then
        Text1.Text = rs.Fields("tblAttendance")
        Text2.Text = rs.Fields("tblDate")
        Text3.Text = rs.Fields("tblShiftPump")
        
        Command1.Enabled = True
    Else
        Text1.Text = rs.Fields("tblAttendance")
        Text2.Text = rs.Fields("tblDate")
        Text3.Text = rs.Fields("tblShiftPump")
        
        Command1.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    If MsgBox("are you sure want to delete this record!", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    '
    ' - if there is no date
    If Text2 = "" Then
        Exit Sub
    End If
    
    If CInt(Text3) <> 0 Then
        MsgBox "Please remove this record from the Sales Delete form!", vbExclamation
        frmDelete.Show
        Unload Me
        Exit Sub
    End If
    sSQL = "delete from tblSalesRecord where tblStaffID = '" & combo1.List(Combo2.ListIndex) & "' and tblDate = #" & Text2.Text & "# and tblAttendance = '" & Text1.Text & "'"
    cn.Execute sSQL
    MsgBox "Entry Deleted!"
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Command1.Enabled = False
End Sub

Private Sub Form_Load()
    sSQL = "select tblStaffID, tblSurname, tblFirstname from tblStaff"
    Set rs = cn.Execute(sSQL)
    combo1.Clear
    Combo2.Clear
    Do While Not rs.EOF
        combo1.AddItem rs.Fields(0)
        Combo2.AddItem rs.Fields(1) & " " & rs.Fields(2)
        rs.MoveNext
    Loop
End Sub
