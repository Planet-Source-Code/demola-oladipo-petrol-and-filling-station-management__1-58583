VERSION 5.00
Begin VB.Form frmPumpSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pump History Report"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmPumpSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Report"
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmPumpSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If List1.ListIndex < 0 Then
        MsgBox "Please select a Pump to view record!", vbCritical
        Exit Sub
    End If

    With dEnv.rscmdPumpSales
        If .State Then
            .Close
            Unload dRptPumpSales
        End If
    End With
    
    dEnv.cmdPumpSales List1.Text
    dRptPumpSales.Show
    dRptPumpSales.Caption = "Information for Pump " & List1.Text
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    sSQL = "select tblPumpID from tblPumpType"
    Set rsPumps = cn.Execute(sSQL)
    With rsPumps
        Do While Not .EOF
            List1.AddItem .Fields("tblPumpID")
            .MoveNext
        Loop
    End With
End Sub

Private Sub List1_DblClick()
    Command1_Click
End Sub

