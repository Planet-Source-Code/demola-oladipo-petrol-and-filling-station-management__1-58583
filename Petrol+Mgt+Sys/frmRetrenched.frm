VERSION 5.00
Begin VB.Form frmRetrenched 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recall Retrenched Staff"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmRetrenched.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4725
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recall"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Staff ID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRetrenched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'    If Combo1.Text = "" Then
'        Exit Sub
'    End If
'    sSQL = "insert into tblStaff from tblRetrench where tblRetrench.tblStaffID = '" & Combo1.Text & "'"
'    cn.Execute sSQL
'    MsgBox Combo1.Text & " has been recalled!", vbInformation
'    loadID
'End Sub
'
'Private Sub Form_Load()
'    loadID
'End Sub
'
'Sub loadID()
'    sSQL = "select tblStaffID from tblRetrench"
'    Set rsStaff = cn.Execute(sSQL)
'    Combo1.Clear
'    With rsStaff
'        Do While Not .EOF
'            Combo1.AddItem .Fields(0)
'            .MoveNext
'        Loop
'    End With
'End Sub

Private Sub Command1_Click()
    If Combo1.Text = "" Then
        Exit Sub
    End If
    
    sSQL = "select * from tblRetrench where tblStaffID = '" & List1.List(Combo1.ListIndex) & "'"
    Set rsStaff = cn.Execute(sSQL)
    
    With rsStaff
        sSQL = "insert into tblStaff(tblStaffID, tblSurname, tblFirstName, tblSex, tblAge, tblAddress, tblPhoneNo, tblNextofKin, tblNOKAddress, tblNOKPhone, tblStaffGuarantor, tblStaffGuarantorAddress, tblStaffGuarantorPhone) values ('" & List1.List(Combo1.ListIndex) & "','" & .Fields("tblSurname") & "','" & .Fields("tblFirstName") & "','" & .Fields("tblSex") & "'," & .Fields("tblAge") & ",'" & .Fields("tblAddress") & "','" & .Fields("tblPhoneNo") & "','" & .Fields("tblNextofKin") & "','" & .Fields("tblNOKAddress") & "','" & .Fields("tblNOKPhone") & "','" & .Fields("tblStaffGuarantor") & "','" & .Fields("tblStaffGuarantorAddress") & "','" & .Fields("tblStaffGuarantorPhone") & "')"
        cn.Execute sSQL
    End With
    
    sSQL = "delete from tblretrench where tblStaffId = '" & List1.List(Combo1.ListIndex) & "'"
    cn.Execute sSQL
    
    frmStaffMgt.LoadStaffID
    loadID
End Sub

Private Sub Form_Load()
    loadID
End Sub

Sub loadID()
    sSQL = "select tblStaffID, tblSurname, tblFirstname from tblRetrench"
    Set rsStaff = cn.Execute(sSQL)
    Combo1.Clear
    List1.Clear
    With rsStaff
        Do While Not .EOF
            List1.AddItem .Fields(0)
            Combo1.AddItem .Fields(1) & " " & .Fields(2)
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStaffMgt.Enabled = True
End Sub

