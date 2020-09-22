VERSION 5.00
Begin VB.Form frmSuspended2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recall - Security Personnel"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmSuspended2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4095
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recall"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Personnel"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSuspended2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Combo1.Text = "" Then
        Exit Sub
    End If
    
    sSQL = "select * from tblSuspend2 where tblStaffID = '" & List1.List(Combo1.ListIndex) & "'"
    Set rsStaff = cn.Execute(sSQL)
    
    With rsStaff
        sSQL = "insert into tblSecurity(tblStaffID, tblSurname, tblFirstName, tblSex, tblAge, tblAddress, tblPhoneNo, tblNextofKin, tblNOKAddress, tblNOKPhone, tblStaffGuarantor, tblStaffGuarantorAddress, tblStaffGuarantorPhone) values ('" & List1.List(Combo1.ListIndex) & "','" & .Fields("tblSurname") & "','" & .Fields("tblFirstName") & "','" & .Fields("tblSex") & "'," & .Fields("tblAge") & ",'" & .Fields("tblAddress") & "','" & .Fields("tblPhoneNo") & "','" & .Fields("tblNextofKin") & "','" & .Fields("tblNOKAddress") & "','" & .Fields("tblNOKPhone") & "','" & .Fields("tblStaffGuarantor") & "','" & .Fields("tblStaffGuarantorAddress") & "','" & .Fields("tblStaffGuarantorPhone") & "')"
        cn.Execute sSQL
    End With
    
    sSQL = "delete from tblSuspend2 where tblStaffId = '" & List1.List(Combo1.ListIndex) & "'"
    cn.Execute sSQL
    
    frmStaffMgt2.LoadStaffID
    loadID
End Sub

Private Sub Form_Load()
    loadID
End Sub

Sub loadID()
    sSQL = "select tblStaffID, tblSurname, tblFirstname from tblSuspend2"
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
    frmStaffMgt2.Enabled = True
End Sub
