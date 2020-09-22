VERSION 5.00
Begin VB.Form frmSelectStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Staff Report"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSelectStaff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4680
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Report"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmSelectStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If List1.ListIndex < 0 Then
        MsgBox "Please select a Staff to view record!", vbCritical
        Exit Sub
    End If

    With dEnv.rscmdSelectStaff
        If .State Then
            .Close
            Unload dRptSelectStaff
        End If
    End With
    dEnv.cmdSelectStaff List2.List(List1.ListIndex)
    dRptSelectStaff.Show
    dRptSelectStaff.Caption = "Information for Staff " & List1.Text
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    sSQL = "select tblStaffID, tblSurname, tblFirstName from tblStaff"
    Set rsStaff = cn.Execute(sSQL)
    With rsStaff
        Do While Not .EOF
            List2.AddItem .Fields("tblStaffID")
            List1.AddItem .Fields("tblSurname") & " " & .Fields("tblFirstName")
            .MoveNext
        Loop
    End With
End Sub

Private Sub List1_DblClick()
    Command1_Click
End Sub
