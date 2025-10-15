Option Explicit

' Excel Function Sync VBA Module
' Replicates the functionality of jsonexcelexctraction.cmd
' Usage: Import this module into Excel, then run ShowFunctionSyncForm

Private Type FormControls
    ExcelPath As String
    JsonPath As String
    IsInsertMode As Boolean
End Type

Private FormData As FormControls

' Main entry point - shows the function sync form
Public Sub ShowFunctionSyncForm()
    Dim userForm As Object
    Set userForm = CreateFunctionSyncForm()
    userForm.Show
End Sub

' Creates and configures the user form
Private Function CreateFunctionSyncForm() As Object
    Dim frm As Object
    Set frm = UserForms.Add("FunctionSyncForm")
    
    With frm
        .Caption = "Excel Function Sync"
        .Width = 450
        .Height = 220
        .StartUpPosition = 1 ' Center on screen
    End With
    
    ' Add controls programmatically
    AddFormControls frm
    Set CreateFunctionSyncForm = frm
End Function

' Adds all controls to the form
Private Sub AddFormControls(frm As Object)
    Dim btnExcel As Object, btnJson As Object, btnRun As Object
    Dim lblExcel As Object, lblJson As Object, lblStatus As Object
    Dim chkMode As Object, progressBar As Object
    
    ' Excel file button
    Set btnExcel = frm.Controls.Add("Forms.CommandButton.1")
    With btnExcel
        .Name = "btnExcel"
        .Caption = "Select Excel File"
        .Left = 10
        .Top = 20
        .Width = 120
        .Height = 25
    End With
    
    ' Excel file label
    Set lblExcel = frm.Controls.Add("Forms.Label.1")
    With lblExcel
        .Name = "lblExcel"
        .Caption = "No Excel file selected"
        .Left = 140
        .Top = 25
        .Width = 280
        .Height = 20
    End With
    
    ' JSON file button
    Set btnJson = frm.Controls.Add("Forms.CommandButton.1")
    With btnJson
        .Name = "btnJson"
        .Caption = "Select JSON File"
        .Left = 10
        .Top = 60
        .Width = 120
        .Height = 25
    End With
    
    ' JSON file label
    Set lblJson = frm.Controls.Add("Forms.Label.1")
    With lblJson
        .Name = "lblJson"
        .Caption = "Default will be auto-generated"
        .Left = 140
        .Top = 65
        .Width = 280
        .Height = 20
    End With
    
    ' Mode checkbox
    Set chkMode = frm.Controls.Add("Forms.CheckBox.1")
    With chkMode
        .Name = "chkMode"
        .Caption = "Insert Mode (unchecked = Extract Mode)"
        .Left = 10
        .Top = 100
        .Width = 250
        .Height = 20
    End With
    
    ' Status label
    Set lblStatus = frm.Controls.Add("Forms.Label.1")
    With lblStatus
        .Name = "lblStatus"
        .Caption = "Idle"
        .Left = 10
        .Top = 130
        .Width = 400
        .Height = 20
    End With
    
    ' Run button
    Set btnRun = frm.Controls.Add("Forms.CommandButton.1")
    With btnRun
        .Name = "btnRun"
        .Caption = "Run"
        .Left = 10
        .Top = 160
        .Width = 80
        .Height = 30
    End With
    
    ' Set up event handlers
    SetupEventHandlers frm
End Sub

' Sets up event handlers for form controls
Private Sub SetupEventHandlers(frm As Object)
    ' Note: VBA UserForm event handling is typically done in the UserForm code module
    ' This is a simplified version - in practice, you'd create a proper UserForm
    ' For now, we'll handle events in the button click procedures
End Sub

' Handles Excel file selection
Public Sub SelectExcelFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            FormData.ExcelPath = .SelectedItems(1)
            
            ' Auto-generate JSON path
            Dim baseName As String, folderPath As String
            baseName = Left(Dir(FormData.ExcelPath), InStrRev(Dir(FormData.ExcelPath), ".") - 1)
            folderPath = Left(FormData.ExcelPath, InStrRev(FormData.ExcelPath, "\"))
            FormData.JsonPath = folderPath & baseName & " - functions.json"
            
            UpdateFormLabels
        End If
    End With
End Sub

' Handles JSON file selection
Public Sub SelectJsonFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select JSON File"
        .Filters.Clear
        .Filters.Add "JSON Files", "*.json"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            FormData.JsonPath = .SelectedItems(1)
            UpdateFormLabels
        End If
    End With
End Sub

' Updates form labels with current file paths
Private Sub UpdateFormLabels()
    ' This would update the actual form labels in a real UserForm implementation
    Debug.Print "Excel: " & FormData.ExcelPath
    Debug.Print "JSON: " & FormData.JsonPath
End Sub

' Main execution function - Extract or Insert mode
Public Sub RunFunctionSync(insertMode As Boolean)
    If FormData.ExcelPath = "" Then
        MsgBox "Please select an Excel file.", vbExclamation
        Exit Sub
    End If
    
    If FormData.JsonPath = "" Then
        MsgBox "Please select or confirm a JSON path.", vbExclamation
        Exit Sub
    End If
    
    FormData.IsInsertMode = insertMode
    
    If insertMode Then
        InsertFunctionsFromJson
    Else
        ExtractFunctionsToJson
    End If
End Sub

' Extract LAMBDA functions from Excel to JSON
Private Sub ExtractFunctionsToJson()
    Dim wb As Workbook
    Dim nm As Name
    Dim functionDict As Object
    Dim jsonText As String
    Dim fileNum As Integer
    
    ' Create dictionary to store functions
    Set functionDict = CreateObject("Scripting.Dictionary")
    
    ' Open the target workbook
    On Error GoTo ErrorHandler
    Set wb = Workbooks.Open(FormData.ExcelPath, ReadOnly:=True)
    
    UpdateStatus "Scanning named ranges..."
    
    ' Scan all named ranges for LAMBDA functions
    For Each nm In wb.Names
        If Left(nm.RefersTo, 7) = "=LAMBDA" Then
            functionDict(nm.Name) = nm.RefersTo
        End If
    Next nm
    
    wb.Close SaveChanges:=False
    
    ' Convert to JSON format
    jsonText = DictionaryToJson(functionDict)
    
    ' Write to file
    fileNum = FreeFile
    Open FormData.JsonPath For Output As #fileNum
    Print #fileNum, jsonText
    Close #fileNum
    
    UpdateStatus "Exported " & functionDict.Count & " functions to " & FormData.JsonPath
    MsgBox "Successfully exported " & functionDict.Count & " LAMBDA functions!", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "Error during extraction: " & Err.Description, vbCritical
End Sub

' Insert LAMBDA functions from JSON to Excel
Private Sub InsertFunctionsFromJson()
    Dim wb As Workbook
    Dim jsonText As String
    Dim functionDict As Object
    Dim key As Variant
    Dim fileNum As Integer
    Dim insertCount As Integer
    
    ' Check if JSON file exists
    If Dir(FormData.JsonPath) = "" Then
        MsgBox "JSON file not found: " & FormData.JsonPath, vbCritical
        Exit Sub
    End If
    
    ' Read JSON file
    fileNum = FreeFile
    Open FormData.JsonPath For Input As #fileNum
    jsonText = Input(LOF(fileNum), #fileNum)
    Close #fileNum
    
    ' Parse JSON (simplified - you might want to use a proper JSON parser)
    Set functionDict = JsonToDictionary(jsonText)
    
    ' Open the target workbook
    On Error GoTo ErrorHandler
    Set wb = Workbooks.Open(FormData.ExcelPath)
    
    UpdateStatus "Inserting functions..."
    insertCount = 0
    
    ' Insert each function as a named range
    For Each key In functionDict.Keys
        On Error Resume Next
        ' Delete existing name if it exists
        wb.Names(CStr(key)).Delete
        On Error GoTo ErrorHandler
        
        ' Add the new named range with LAMBDA formula
        wb.Names.Add Name:=CStr(key), RefersTo:=functionDict(key)
        insertCount = insertCount + 1
        
        UpdateStatus "Inserted " & insertCount & " of " & functionDict.Count & " functions..."
    Next key
    
    wb.Save
    wb.Close SaveChanges:=True
    
    UpdateStatus "Successfully inserted " & insertCount & " functions!"
    MsgBox "Successfully inserted " & insertCount & " LAMBDA functions!", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "Error during insertion: " & Err.Description, vbCritical
End Sub

' Updates status (in a real form, this would update a label)
Private Sub UpdateStatus(message As String)
    Debug.Print "Status: " & message
    DoEvents ' Allow UI to refresh
End Sub

' Simple JSON to Dictionary converter (basic implementation)
Private Function JsonToDictionary(jsonText As String) As Object
    Dim dict As Object
    Dim lines As Variant
    Dim i As Integer
    Dim line As String
    Dim colonPos As Integer
    Dim key As String, value As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Remove braces and split by lines
    jsonText = Replace(jsonText, "{", "")
    jsonText = Replace(jsonText, "}", "")
    jsonText = Replace(jsonText, Chr(13), "")
    lines = Split(jsonText, Chr(10))
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If line <> "" And line <> "," Then
            ' Remove trailing comma
            If Right(line, 1) = "," Then
                line = Left(line, Len(line) - 1)
            End If
            
            colonPos = InStr(line, ":")
            If colonPos > 0 Then
                key = Trim(Mid(line, 2, colonPos - 3)) ' Remove quotes
                value = Trim(Mid(line, colonPos + 1))
                value = Mid(value, 2, Len(value) - 2) ' Remove quotes
                
                ' Unescape JSON characters
                value = Replace(value, "\""", """")
                value = Replace(value, "\\", "\")
                
                dict(key) = value
            End If
        End If
    Next i
    
    Set JsonToDictionary = dict
End Function

' Simple Dictionary to JSON converter
Private Function DictionaryToJson(dict As Object) As String
    Dim result As String
    Dim key As Variant
    Dim value As String
    Dim firstItem As Boolean
    
    result = "{" & vbCrLf
    firstItem = True
    
    For Each key In dict.Keys
        If Not firstItem Then
            result = result & "," & vbCrLf
        End If
        
        value = CStr(dict(key))
        ' Escape JSON characters
        value = Replace(value, "\", "\\")
        value = Replace(value, """", "\""")
        
        result = result & "  """ & CStr(key) & """: """ & value & """"
        firstItem = False
    Next key
    
    result = result & vbCrLf & "}"
    DictionaryToJson = result
End Function

' Utility function to get current workbook path for default Excel selection
Public Function GetCurrentWorkbookPath() As String
    If Not ActiveWorkbook Is Nothing Then
        GetCurrentWorkbookPath = ActiveWorkbook.FullName
    Else
        GetCurrentWorkbookPath = ""
    End If
End Function

' Simple form interface using InputBox and MsgBox (alternative to UserForm)
Public Sub SimpleFunctionSync()
    Dim excelPath As String, jsonPath As String
    Dim mode As Integer
    Dim baseName As String, folderPath As String
    
    ' Get Excel file path
    excelPath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", , "Select Excel File")
    If excelPath = "False" Then Exit Sub
    
    ' Auto-generate JSON path
    baseName = Left(Dir(excelPath), InStrRev(Dir(excelPath), ".") - 1)
    folderPath = Left(excelPath, InStrRev(excelPath, "\"))
    jsonPath = folderPath & baseName & " - functions.json"
    
    ' Ask for JSON path confirmation
    jsonPath = InputBox("JSON file path:", "Confirm JSON Path", jsonPath)
    If jsonPath = "" Then Exit Sub
    
    ' Ask for mode
    mode = MsgBox("Choose mode:" & vbCrLf & _
                  "YES = Insert functions from JSON to Excel" & vbCrLf & _
                  "NO = Extract functions from Excel to JSON", _
                  vbYesNoCancel + vbQuestion, "Select Mode")
    
    If mode = vbCancel Then Exit Sub
    
    ' Set form data
    FormData.ExcelPath = excelPath
    FormData.JsonPath = jsonPath
    
    ' Run the sync
    If mode = vbYes Then
        InsertFunctionsFromJson
    Else
        ExtractFunctionsToJson
    End If
End Sub