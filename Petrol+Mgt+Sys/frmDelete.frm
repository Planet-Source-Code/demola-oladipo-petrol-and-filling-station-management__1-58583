VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDelete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Delete"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6015
   Begin VB.ComboBox dataCombo1 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   22
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sales Record Details"
      Enabled         =   0   'False
      Height          =   2295
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   5415
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label10 
         Caption         =   "Shift"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Pump Number"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Staff Name"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pump Record Detail"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Final Reading"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Initial Reading"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Date"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Pump Number"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6960
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Petrol\Mobil\db\dbase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Petrol\Mobil\db\dbase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblPumpType"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Last Reading"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Select Pump"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim initial As Double

Private Sub Command1_Click()
    If Text1.Text = "" Then Exit Sub
    If dataCombo1.Text = "" Then Exit Sub
    
    'ask to sort desc
    sSQL = "select * from tblPumpRecord where tblPumpID = " & CInt(dataCombo1.Text) & " and tblFinalLitres = " & CDbl(Text1.Text)
    Set rs = cn.Execute(sSQL)
    If rs.EOF And rs.BOF Then
        Exit Sub
    End If
    
    initial = rs.Fields("tblInitialLitres")
    
    sSQL = "select * from tblSalesRecord where tblShiftPump = " & rs.Fields("tblPumpID") & " and tblShift = '" & rs.Fields("tblShift") & "' and tblStaffID = '" & rs.Fields("tblStaffID") & "'"
    'MsgBox sSQL
    Set rs1 = cn.Execute(sSQL)
    If rs1.EOF And rs1.BOF Then
        Exit Sub
    End If
    
    'MsgBox rs.Fields("tblStaffID") & ":" & rs.Fields("tblPumpID") & ":" & rs.Fields("tblLitresSold")
    'MsgBox rs1.Fields("tblStaffID") & ":" & rs1.Fields("tblshiftPump") & ":" & rs1.Fields("tblShiftLitres")
    Text2 = rs.Fields("tblPumpID")
    Text3 = rs.Fields("tblDate")
    Text4 = rs.Fields("tblInitialLitres")
    Text5 = rs.Fields("tblFinalLitres")
    
    Text6 = rs1.Fields("tblShiftPump")
    Text7 = rs1.Fields("tblDate")
    Text8 = rs.Fields("tblStaffName")
    Text9 = rs1.Fields("tblShift")
    Text10 = rs.Fields("tblStaffID")
    
End Sub

Private Sub Command2_Click()
    If MsgBox("are you sure you want to delete this record?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

On Error GoTo roll
    'cn.BeginTrans
    sSQL = "update tblPumpType set tblLastReading = " & initial & " where tblPumpId = " & dataCombo1.Text
    cn.Execute sSQL

    sSQL = "delete from tblPumpRecord where tblPumpID = " & CInt(dataCombo1.Text) & " and tblFinalLitres = " & CDbl(Text1.Text)
    cn.Execute sSQL
    
    'sSQL = "select * from tblSalesRecord where tblSalesRecord.tblShiftPump = " & CInt(Text2) & " and tblSalesRecord.tblShift = '" & Text9 & "' and tblSalesRecord.tblStaffID = '" & Text10 & "' and tblSalesRecord.tblDate = #" & Text3.Text & "#"
    
    sSQL = "delete from tblSalesRecord where tblSalesRecord.tblShiftPump = " & CInt(Text2) & " and tblSalesRecord.tblShift = '" & Text9 & "' and tblSalesRecord.tblStaffID = '" & Text10 & "' and tblSalesRecord.tblDate = #" & Text3.Text & "#"
    
    cn.Execute sSQL
    
    'cn.CommitTrans
    
    clearText
    DataCombo1_Click
    'Command1_Click
    Exit Sub
    
roll:
   ' cn.RollbackTrans
   MsgBox Err.Description
End Sub

Private Sub DataCombo1_Click()
    sSQL = "select tblLastReading from tblPumpType where tblPumpID = " & dataCombo1.Text
    Set rs = cn.Execute(sSQL)
    clearText
    Text1 = rs.Fields(0)
    rs.Close
End Sub

Private Sub Form_Load()
sSQL = "select tblPumpid from tblpumptype"
Set rs = cn.Execute(sSQL)
dataCombo1.Clear
Do While Not rs.EOF
    dataCombo1.AddItem rs.Fields(0)
    rs.MoveNext
Loop
End Sub

Sub clearText()
Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Text10 = ""
End Sub
