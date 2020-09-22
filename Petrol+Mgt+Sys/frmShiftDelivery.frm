VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShiftDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift Delivery Report"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "View"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59441153
      CurrentDate     =   37917
   End
   Begin VB.Label Label2 
      Caption         =   "Shift"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmShiftDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   If Combo1.Text = "" Then
        MsgBox "Please select a Shift to view!", vbCritical
        Exit Sub
    End If

    With dEnv.rscmdShiftDelivery
        If .State Then
            .Close
            Unload dRPtSD
        End If
    End With
    Dim cs As String
    cs = "'" & Format(DTPicker1.Value, "dd mmm yy") & "'"
    dEnv.cmdShiftDelivery Combo1.Text
    dRPtSD.Show
    'dRptSd.Caption = "Information for Staff " & List1.Text
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Morning"
    Combo1.AddItem "Night"
    
End Sub
