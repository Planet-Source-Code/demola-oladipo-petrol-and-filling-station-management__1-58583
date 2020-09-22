Attribute VB_Name = "Module1"
'Public strStaffID As String
Global connStr, strProdName, dBaseName As String

Global cn As New Connection
Global cn1 As New Connection
Global rs As New Recordset
Global rsStaff As New Recordset
Global rsPumps As New Recordset
Global rsLogin As New Recordset

Global StartingAddress As String
Global LoginSucceeded As Boolean
Global PumpLitreShown As Boolean
Global appStart As Boolean
Sub main()
    On Error GoTo ErrorHandler
    'display splash screen
    With frmSplash
        .Show
        .Refresh
    End With
    
    appStart = True
    'initialize variables
    strProdName = "Mobil Service Station"
    
    dBaseName = "dbase.mdb"
    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\" & dBaseName & ";Persist Security Info=False"
    cn.Open connStr
    
    dBaseName = "dump.mdb"
    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\" & dBaseName & ";Persist Security Info=False"
    cn1.Open connStr
    
    Exit Sub

ErrorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "System Files not found! Aborting!", vbCritical, "Re-Install Application"
        End
    ElseIf Err.Number = 3705 Then
        MsgBox "Unclosed System Files!", vbCritical, "Aborting"
        End
    Else
        MsgBox Err.Number & Err.Description
    End If
End Sub

Sub disableMainMenu(enableMode As Boolean)
    With frmMain
        .mnuFile.Enabled = enableMode
        .mnuHelp.Enabled = enableMode
        .mnuOperation.Enabled = enableMode
        .mnuMgt.Enabled = enableMode
        .mnuAdmin.Enabled = enableMode
        .mnuWindow.Enabled = enableMode
        .mnuView.Enabled = enableMode
        .mnuReports.Enabled = enableMode
    End With
End Sub

Public Function LoadRecordSetIntoGrid(ctlGrid As MSFlexGrid, _
   rs As adodb.Recordset) As Boolean

    Dim sTmp As String
    Dim vArr As Variant
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'first we have to clear the grid from previous query results
    '**** remove existing grid rows
    ctlGrid.Rows = 2
    ctlGrid.AddItem ""
    ctlGrid.RemoveItem 1
    
    If Not rs.EOF And Not rs.BOF Then
        'get the recordset into a string delimiting the fields
        'with a tab character and the records with a semicolon
        'replace nulls with a space
        sTmp = rs.GetString(adClipString, , Chr(9), ";", " ")

        'split the string into an array of individual records
        vArr = Split(sTmp, ";")
        
        'now add the records to the grid
        For i = 0 To UBound(vArr) - 1
            ctlGrid.AddItem vArr(i)
        Next
        'set the return value
        LoadRecordSetIntoGrid = True
    Else
        LoadRecordSetIntoGrid = False
        Exit Function
    End If
    
    'remove empty rows
    If ctlGrid.Rows > 2 Then
        On Error Resume Next 'get over removing a single record
         For i = 1 To ctlGrid.Rows - 1
            If ctlGrid.TextMatrix(i, 1) = "" Then
                ctlGrid.RemoveItem i
            Else
                Exit For
            End If
        Next
    End If

    Exit Function
ErrorHandler:
    
    LoadRecordSetIntoGrid = False
End Function

