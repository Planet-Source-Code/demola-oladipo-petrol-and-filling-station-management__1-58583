VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStaffMgt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Management - Sales Personnel"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   Icon            =   "frmStaffMgt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9720
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Caption         =   "Picture"
      Enabled         =   0   'False
      Height          =   3735
      Left            =   4440
      TabIndex        =   36
      Top             =   240
      Width           =   3015
      Begin VB.CommandButton Command7 
         Caption         =   "Locate Image"
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   2295
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Management"
      Height          =   3735
      Left            =   7680
      TabIndex        =   31
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "Employ Staff"
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit Staff Information"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Retrench Staff"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Suspend Staff"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   32
         Top             =   2880
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Recall"
      Height          =   2175
      Left            =   7680
      TabIndex        =   28
      Top             =   3960
      Width           =   1935
      Begin VB.CommandButton Command6 
         Caption         =   "Retrenched Staff"
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Suspended Staff"
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Next of Kin"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   3615
      Begin VB.TextBox txtNOKPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtNOKAddress 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNOK 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Phone Number"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Staff Gurantor"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   3840
      TabIndex        =   2
      Top             =   4080
      Width           =   3615
      Begin VB.TextBox txtStaffGPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtStaffGAddress 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtStaffG 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Phone Number"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Personal Details"
      Enabled         =   0   'False
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
      Begin VB.TextBox txtPhoneNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox cboSex 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmStaffMgt.frx":030A
         Left            =   1200
         List            =   "frmStaffMgt.frx":0314
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtSurname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "Phone Number"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Age"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Sex"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Surname"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Staff ID Numbers"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox list1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   360
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmStaffMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New Recordset

Dim bEdited As Boolean
Dim sSQL, staffID  As String

Private Sub Command1_Click()
    
    If Command1.Caption <> "Save Record" Then
        Frame1.Enabled = False
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        Frame7.Enabled = True
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        ClearFields
        Me.txtSurname.SetFocus
        Command1.Caption = "Save Record"
    Else
        Command1.Caption = "Employ Staff"
        Frame1.Enabled = True
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        Frame7.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        'create SQL stmt
        If checkFields Then
            staffID = InputBox("Please assign to this Staff an ID Number!" & vbCrLf & "Format (CS-2003-001)")
            sSQL = "select tblStaffID from tblStaff where tblStaffID = '" & staffID & "'"
            Set rsTemp = cn.Execute(sSQL)
            If rsTemp.BOF And rsTemp.EOF Then
                sSQL = "insert into tblStaff(tblStaffID, tblSurname, tblFirstName, tblSex, tblAge, tblAddress, tblPhoneNo, tblNextofKin, tblNOKAddress, tblNOKPhone, tblStaffGuarantor, tblStaffGuarantorAddress, tblStaffGuarantorPhone) values ('" & UCase$(staffID) & "','" & txtSurname & "','" & txtFirstName & "','" & cboSex & "'," & txtAge & ",'" & Me.txtAddress & "','" & Me.txtPhoneNo & "','" & Me.txtNOK & "','" & Me.txtNOKAddress & "','" & Me.txtNOKPhone & "','" & Me.txtStaffG & "','" & Me.txtStaffGAddress & "','" & Me.txtStaffGPhone & "')"
                
                'MsgBox sSQL
                cn.Execute sSQL
            Else
                MsgBox "Duplicate ID found. Please assign another staff number!"
                Exit Sub
            End If
            
            
            'this saves the users picture using the staffid as filename
            On Error Resume Next
            SavePicture Image1.Picture, App.Path & "\pictures\" & staffID & ".jpg"
            
        End If
        LoadStaffID
        ClearFields
    End If
End Sub


Public Sub LoadStaffID()
    sSQL = "select tblStaffID from tblStaff"
    Set rsStaff = cn.Execute(sSQL)
    With rsStaff
        list1.Clear
        Do While Not .EOF
            list1.AddItem .Fields("tblStaffID")
            .MoveNext
        Loop
    End With
End Sub

Public Sub Command2_Click()
    If Command2.Caption <> "Save Record" Then
        Command1.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command2.Caption = "Save Record"
        Frame1.Enabled = False
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        Frame7.Enabled = True
        Frame5.Enabled = False
    Else
        If checkFields = False Then
            Exit Sub
        End If
        Command2.Caption = "Edit Staff Information"
        'sSQL = "update tblStaff(tblStaffID, tblSurname, tblFirstName, tblSex, tblAge, tblAddress, tblPhoneNo, tblNextofKin, tblNOKAddress, tblNOKPhone, tblStaffGuarantor, tblStaffGuarantorAddress, tblStaffGuarantorPhone) values ('" & staffID & "','" & txtSurname & "','" & txtFirstName & "','" & cboSex & "'," & txtAge & ",'" & Me.txtAddress & "','" & Me.txtPhoneNo & "','" & Me.txtNOK & "','" & Me.txtNOKAddress & "','" & Me.txtNOKPhone & "','" & Me.txtStaffG & "','" & Me.txtStaffGAddress & "','" & Me.txtStaffGPhone & "') where tblStaffId = '" & List1.Text & "'"
        sSQL = "update tblStaff set tblSurname = '" & txtSurname & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblFirstName = '" & txtFirstName & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblAge = '" & txtAge & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblAddress = '" & txtAddress & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblPhoneNo = '" & txtPhoneNo & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblNextofKin = '" & txtNOK & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblNOKAddress = '" & txtNOKAddress & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblNOKPhone = '" & txtNOKPhone & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblStaffGuarantor = '" & txtStaffG & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblStaffGuarantorAddress = '" & txtStaffGAddress & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        sSQL = "update tblStaff set tblStaffGuarantorPhone = '" & txtStaffGPhone & "' where tblSTaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        SavePicture Image1.Picture, App.Path & "\pictures\" & list1.Text & ".jpg"
        Command1.Enabled = True
        Command3.Enabled = True
        Frame5.Enabled = True
        Frame1.Enabled = True
        LoadStaffID
        Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Frame2.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        Frame7.Enabled = False
    End If
End Sub

Private Sub Command3_Click()
    If list1.ListIndex <> -1 Then
        sSQL = "insert into tblRetrench(tblStaffID, tblSurname, tblFirstName, tblSex, tblAge, tblAddress, tblPhoneNo, tblNextofKin, tblNOKAddress, tblNOKPhone, tblStaffGuarantor, tblStaffGuarantorAddress, tblStaffGuarantorPhone) values ('" & list1.Text & "','" & txtSurname & "','" & txtFirstName & "','" & cboSex & "'," & txtAge & ",'" & Me.txtAddress & "','" & Me.txtPhoneNo & "','" & Me.txtNOK & "','" & Me.txtNOKAddress & "','" & Me.txtNOKPhone & "','" & Me.txtStaffG & "','" & Me.txtStaffGAddress & "','" & Me.txtStaffGPhone & "')"
        cn.Execute sSQL
        
        sSQL = "delete from tblStaff where tblStaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        ClearFields
        LoadStaffID
    End If
End Sub

Private Sub Command4_Click()
    If list1.ListIndex <> -1 Then
        sSQL = "insert into tblSuspend(tblStaffID, tblSurname, tblFirstName, tblSex, tblAge, tblAddress, tblPhoneNo, tblNextofKin, tblNOKAddress, tblNOKPhone, tblStaffGuarantor, tblStaffGuarantorAddress, tblStaffGuarantorPhone) values ('" & list1.Text & "','" & txtSurname & "','" & txtFirstName & "','" & cboSex & "'," & txtAge & ",'" & Me.txtAddress & "','" & Me.txtPhoneNo & "','" & Me.txtNOK & "','" & Me.txtNOKAddress & "','" & Me.txtNOKPhone & "','" & Me.txtStaffG & "','" & Me.txtStaffGAddress & "','" & Me.txtStaffGPhone & "')"
        cn.Execute sSQL
        
        sSQL = "delete from tblStaff where tblStaffId = '" & list1.Text & "'"
        cn.Execute sSQL
        ClearFields
        LoadStaffID
    End If
End Sub

Private Sub Command5_Click()
    frmSuspended.Show
    Me.Enabled = False
End Sub

Private Sub Command6_Click()
    frmRetrenched.Show
End Sub

Private Sub Command7_Click()
    With cDlg
        .InitDir = App.Path & "\pictures\"
        .Filter = "Jpeg Photographs|*.jpg|Windows Bitmaps|*.bmp|All Files|*.*"
        .DialogTitle = "Load Staff Image File"
        .FileName = ""
        Image1.Picture = LoadPicture("")
        Do While .FileName = ""
            .ShowOpen
            If .FileName = "" Then
                If MsgBox("Are you sure you dont want to add a picture?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    GoTo LoopOut
                End If
            End If
        Loop
        Image1.Picture = LoadPicture(.FileName)
LoopOut:
        Command2.Caption = "Save Record"
        ' -- debug
        .FileName = ""
        ' -- end debug
        
        'Me.Command2_Click
    End With
End Sub

Private Sub Form_Load()
    LoadStaffID
End Sub

Private Sub List1_Click()
    On Error Resume Next
    sSQL = "select * from tblstaff where tblstaffid = '" & list1.Text & "'"
    Set rsStaff = cn.Execute(sSQL)
    
    If rsStaff.BOF And rsStaff.EOF Then
        MsgBox "This is an error situation!"
        Exit Sub
    End If
    ClearFields
    With rsStaff
        txtSurname = .Fields("tblSurname")
        txtFirstName = .Fields("tblFirstname")
        cboSex.Text = .Fields("tblSex")
        txtAge = .Fields("tblage")
        txtAddress = .Fields("tblAddress")
        txtPhoneNo = .Fields("tblPhoneNo")
        Me.txtNOK = .Fields("tblNextofKin")
        Me.txtNOKAddress = .Fields("tblNOKAddress")
        Me.txtNOKPhone = .Fields("tblNOKPhone")
        Me.txtStaffG = .Fields("tblStaffGuarantor")
        Me.txtStaffGPhone = .Fields("tblStaffGuarantorPhone")
        Me.txtStaffGAddress = .Fields("tblStaffGuarantorAddress")
        Image1.Picture = LoadPicture(App.Path & "\pictures\" & .Fields("tblStaffID") & ".jpg")
    End With
    Frame7.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    
End Sub

Sub ClearFields()
    txtSurname = ""
    txtFirstName = ""
    cboSex = ""
    txtAge = ""
    txtAddress = ""
    txtPhoneNo = ""
    
    txtStaffG = ""
    txtStaffGAddress = ""
    txtStaffGPhone = ""
    
    txtNOK = ""
    txtNOKAddress = ""
    txtNOKPhone = ""
    
    Image1.Picture = Nothing
End Sub

Function checkFields() As Boolean
    If Me.txtSurname = "" Then
        MsgBox "Please enter a Surname!", vbCritical
        Me.txtSurname.SetFocus
        Exit Function
    End If
    
    If Me.txtFirstName = "" Then
        MsgBox "Please enter a First Name!", vbCritical
        Me.txtFirstName.SetFocus
        Exit Function
    End If
    
    If Me.cboSex = "" Then
        MsgBox "Please select a Sex type!", vbCritical
        Me.cboSex.SetFocus
        Exit Function
    End If
    
    If Me.txtAge = "" Then
        MsgBox "Please enter an Age value!", vbCritical
        Me.txtAge.SetFocus
        Exit Function
    End If
    
    Dim i As Integer
    On Error GoTo Ouch
    i = CInt(txtAge)
    On Error GoTo 0
    
    If Me.txtAddress = "" Then
        MsgBox "Please enter an Address!", vbCritical
        Me.txtAddress.SetFocus
        Exit Function
    End If
    
    If Me.txtPhoneNo = "" Then
        MsgBox "Please enter a Phone contact!", vbCritical
        Me.txtPhoneNo.SetFocus
        Exit Function
    End If
    
    If Me.txtNOK = "" Then
        MsgBox "Please enter a Next of Kin!", vbCritical
        Me.txtNOK.SetFocus
        Exit Function
    End If
    
    If Me.txtNOKAddress = "" Then
        MsgBox "Please enter a Next of Kin address!", vbCritical
        Me.txtNOKAddress.SetFocus
        Exit Function
    End If
    
    If Me.txtNOKPhone = "" Then
        MsgBox "Please enter a Next of Kin Phone Contact!", vbCritical
        Me.txtNOKPhone.SetFocus
        Exit Function
    End If
    
    If Me.txtStaffG = "" Then
        MsgBox "Please enter a Staff Gurantor!", vbCritical
        Me.txtStaffG.SetFocus
        Exit Function
    End If
    
    If Me.txtStaffGAddress = "" Then
        MsgBox "Please enter a Staff Gurantor Address!", vbCritical
        Me.txtStaffGAddress.SetFocus
        Exit Function
    End If
    
    If Me.txtStaffGPhone = "" Then
        MsgBox "Please enter a Staff Gurantors Phone Contact!", vbCritical
        Me.txtStaffGPhone.SetFocus
        Exit Function
    End If
    
    'If Image1. = Nothing Then
    '    MsgBox "Please Staff Picture!", vbCritical
    '    Exit Function
    'End If
    checkFields = True
    
    Exit Function
    
Ouch:
    MsgBox "Please enter an Numeric Age value!", vbCritical
    Me.txtAge.SetFocus
        
End Function

