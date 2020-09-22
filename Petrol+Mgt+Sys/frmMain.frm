VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mobil Service Station"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8580
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Picture         =   "frmMain.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1620
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   2858
      ButtonWidth     =   2487
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Operations"
            Key             =   "operations"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sales"
                  Text            =   "Sales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "delete"
                  Text            =   "Remove Pump Record"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "remove"
                  Text            =   "Remove Staff Entry"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Product Cost"
            Key             =   "pcost"
            Object.Tag             =   "Product Cost"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pumps Mgt"
            Key             =   "pumps"
            Object.ToolTipText     =   "Pumps Management"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bonus"
            Key             =   "bonus"
            Object.ToolTipText     =   "Bonus Computatio for Male Staff"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bonus1"
                  Text            =   "Male Staff"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bonus2"
                  Text            =   "Female Staff"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Staff Management"
            Key             =   "staff"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ccare"
                  Text            =   "Customer Care"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sstaff"
                  Text            =   "Senior Staff"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "security"
                  Text            =   "Security"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Key             =   "report"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   10
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "allstaff"
                  Text            =   "All Staff"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ssalary"
                  Text            =   "Staff Salary"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "fread"
                  Text            =   "Pumps (Final Reading)"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "allsales"
                  Text            =   "Pumps (All Sales)"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "shift"
                  Text            =   "Shift Delivery Report"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pdaily"
                  Text            =   "All Pumps Daily Report"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cShift"
                  Text            =   "Cummulative Shift Delivery"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "about"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6015
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4445
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "12/18/03"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "7:23 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOperation 
      Caption         =   "&Operations"
      Begin VB.Menu mnuSalesRec 
         Caption         =   "&Sales Record"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Pump Record"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove Staff Record"
      End
   End
   Begin VB.Menu mnuMgt 
      Caption         =   "&Management"
      Begin VB.Menu mnuProdMgt 
         Caption         =   "&Product Cost Mgt"
      End
      Begin VB.Menu mnuPumpsMgt 
         Caption         =   "Pumps Management"
      End
      Begin VB.Menu mnuHyp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBonus 
         Caption         =   "&Bonus"
         Begin VB.Menu mnuFemale 
            Caption         =   "&Female"
         End
         Begin VB.Menu mnuMale 
            Caption         =   "&Male"
         End
      End
      Begin VB.Menu mnuPersonnel 
         Caption         =   "Customer Care"
      End
      Begin VB.Menu mnuSeniorStaff 
         Caption         =   "&Senior Staff"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "S&ecurity"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuAllStaff 
         Caption         =   "All Staff Sales"
      End
      Begin VB.Menu mnuSelectStaff 
         Caption         =   "Select Staff"
      End
      Begin VB.Menu mnuHyp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalaryStructure 
         Caption         =   "Staff Salary Report"
      End
      Begin VB.Menu mnuHyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepPumps 
         Caption         =   "Pumps (Final Reading)"
      End
      Begin VB.Menu mnuPumpSales 
         Caption         =   "Pumps (All Sales)"
      End
      Begin VB.Menu mnuHyp4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepSD 
         Caption         =   "Shift Delivery Report"
      End
      Begin VB.Menu mnuCSDR 
         Caption         =   "Cummulative Shift Delivery Report"
      End
      Begin VB.Menu mnuPDReport 
         Caption         =   "All Pumps Daily Report"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Admin"
      Begin VB.Menu mnuPassword 
         Caption         =   "&Change Password"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascde 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuManual 
         Caption         =   "Manual"
      End
      Begin VB.Menu mnuHyp5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Activate()
MDIForm_Load
End Sub

Private Sub MDIForm_Load()

On Error GoTo Ouch

'    If GetStringSetting(strProdName, "Config", "Product Cost", "False") = "False" Then
'        Me.mnuOperation.Enabled = False
'    Else
'        Me.mnuOperation.Enabled = True
'    End If
'
'    If GetStringSetting(strProdName, "Config", "Pump Management", "False") = "False" Then
'        Me.mnuOperation.Enabled = False
'    Else
'        Me.mnuOperation.Enabled = True
'    End If
'
'    If GetStringSetting(strProdName, "Config", "Male Bonus", "False") = "False" Then
'        Me.mnuOperation.Enabled = False
'    Else
'        Me.mnuOperation.Enabled = True
'    End If
'
'    If GetStringSetting(strProdName, "Config", "Female Bonus", "False") = "False" Then
'        Me.mnuOperation.Enabled = False
'    Else
'        Me.mnuOperation.Enabled = True
'    End If
    
    'set printer orientation to landscape
    'Printer.Orientation = vbPRORLandscape

Exit Sub

Ouch:
    If Err.Number = 482 Then
        MsgBox "This package requires that you have a printer installed", vbCritical
        End
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        MsgBox "Please exit from the File Menu", vbCritical
        Cancel = True
    End If
End Sub


Public Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Public Sub mnuAllStaff_Click()
    With dEnv.rscmdAllStaff
        If .State Then
            .Close
            Unload dRptAllStaff
        End If
    End With
    dRptAllStaff.Show
End Sub

Public Sub mnuCascde_Click()
    Me.Arrange vbCascade
End Sub

Public Sub mnuCSDR_Click()
    frmShiftDelivery1.Show
End Sub

Public Sub mnuDelete_Click()
frmDelete.Show
End Sub

Public Sub mnuExit_Click()
    If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Unload Me
    End If
End Sub

Public Sub mnuFemale_Click()
    frmBonus.Show
End Sub

Public Sub mnuMale_Click()
    frmBonus1.Show
End Sub

Public Sub mnuPassword_Click()
    frmLogin1.Show 1
End Sub

Public Sub mnuPDReport_Click()
    frmPDReport.Show
End Sub

Public Sub mnuPersonnel_Click()
    frmStaffMgt.Show
End Sub

Public Sub mnuProdMgt_Click()
    frmProdMgt.Show
End Sub

Public Sub mnuPumpSales_Click()
    frmPumpSales.Show
End Sub

Public Sub mnuPumpsMgt_Click()
    frmPumpMgt.Show
End Sub

Public Sub mnuRemove_Click()
    frmDelete1.Show
End Sub

Public Sub mnuRepPumps_Click()
    With dEnv.rscmdAllPumps
        If .State Then
            .Close
            Unload dRptAllPumps
        End If
    End With
   
    dRptAllPumps.Show
End Sub

Public Sub mnuRepSD_Click()
    frmShiftDelivery.Show
End Sub

Public Sub mnuSalaryStructure_Click()
    frmSalary.Show
End Sub

Public Sub mnuSalesRec_Click()
    frmSales.Show
End Sub

Public Sub mnuSecurity_Click()
    frmStaffMgt2.Show
End Sub

Public Sub mnuSelectStaff_Click()
    frmSelectStaff.Show
End Sub

Public Sub mnuSeniorStaff_Click()
    frmStaffMgt1.Show
End Sub

Public Sub mnuStatus_Click()
    If mnuStatus.Checked Then
        StatusBar1.Visible = False
        mnuStatus.Checked = False
    Else
        StatusBar1.Visible = True
        mnuStatus.Checked = True
    End If
End Sub

Private Sub mnuToolbar_Click()
    If mnuToolbar.Checked Then
        Toolbar1.Visible = False
        mnuToolbar.Checked = False
    Else
        Toolbar1.Visible = True
        mnuToolbar.Checked = True
    End If
End Sub

Private Sub Picture1_Resize()
    With Image1
        .Left = Picture1.Width - .Width - 100
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "pcost"
            Me.mnuProdMgt_Click
        Case "pumps"
            Me.mnuPumpsMgt_Click
        Case "about"
            Me.mnuAbout_Click
        Case "exit"
            Me.mnuExit_Click
        Case Else
            'MsgBox "Error!", vbExclamation
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "bonus1"
            Me.mnuMale_Click
        Case "bonus2"
            Me.mnuFemale_Click
        Case "ccare"
            Me.mnuPersonnel_Click
        Case "sstaff"
            Me.mnuSeniorStaff_Click
        Case "security"
            Me.mnuSecurity_Click
        Case "allstaff"
            Me.mnuAllStaff_Click
        Case "ssalary"
            Me.mnuSalaryStructure_Click
        Case "fread"
            Me.mnuRepPumps_Click
        Case "allsales"
            Me.mnuPumpSales_Click
        Case "shift"
            Me.mnuRepSD_Click
        Case "pdaily"
            Me.mnuPDReport_Click
        Case "cShift"
            Me.mnuCSDR_Click
        Case "sales"
            Me.mnuSalesRec_Click
        Case "delete"
            Me.mnuDelete_Click
        Case "remove"
            Me.mnuRemove_Click
    End Select
End Sub
