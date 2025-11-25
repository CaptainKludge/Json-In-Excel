# 2>nul & @echo off & PowerShell -ExecutionPolicy Bypass -WindowStyle hidden -Command "$s = Get-Content '%~dpnx0' | Select-Object -Skip 1 | Out-String; $s='Set-Location ""%~dp0"";'+$s; $sourceDir = '%~1'; Invoke-Expression $s" & goto :eof

<#
===============================================================================================
|                        EXCEL FUNCTION IMPORT/EXPORT TOOL (JSON)                            |
|                                                                                             |
| DUAL-SCRIPT ARCHITECTURE: CMD wrapper that executes embedded PowerShell code              |
| Part of Brad's PowerShell Tools suite - Windows Context Menu Integration                   |
===============================================================================================

PURPOSE:
Bidirectional Excel function management tool that extracts LAMBDA functions from Excel
workbooks to JSON files for backup/sharing, or imports JSON function definitions back 
into Excel workbooks. Specializes in Excel's advanced LAMBDA function handling.

EXCEL LAMBDA FUNCTIONS:
LAMBDA functions are Excel's custom function feature that allows users to create reusable
formulas with parameters. This tool specifically targets these advanced functions for:
- Backup and version control of custom Excel functions
- Sharing function libraries between workbooks and users
- Bulk function management and deployment

OPERATION MODES:
1. EXTRACT MODE (Default): Scans Excel workbook for LAMBDA functions and exports to JSON
   - Reads all Named Ranges in the workbook
   - Filters for formulas beginning with "=LAMBDA"
   - Creates structured JSON file with function name/formula pairs
   
2. INSERT MODE (Checkbox): Imports function definitions from JSON into Excel workbook
   - Reads JSON file containing function definitions
   - Deletes existing functions with same names (prevents duplicates)
   - Adds new Named Ranges with LAMBDA formulas
   - Saves workbook with updated functions

ARCHITECTURE PATTERN:
This file uses the "dual-script pattern" where:
1. CMD header (line 1) invokes PowerShell with -ExecutionPolicy Bypass
2. PowerShell code starts after this comment block (Skip 1 line)
3. All logic is implemented in PowerShell for rich COM object interaction

INSTALLATION & CONTEXT MENU:
- Install via BradsToolsInstall.reg to add right-click context menu
- Appears on Excel files (.xls, .xlsx, .xlsm) as "Function Insert/Extract"
- Non-background operation - shows GUI for user interaction

EXCEL COM INTEGRATION:
- Uses Excel COM objects for direct workbook manipulation
- Handles Named Ranges and LAMBDA function detection
- Includes proper COM object cleanup to prevent Excel process leaks
- Progress tracking for large function sets

SAFETY FEATURES:
- Automatic JSON path generation based on Excel filename
- Insert mode preserves existing workbook data (only affects Named Ranges)
- Progress feedback during long operations
- Proper error handling for COM object failures

.PARAMETER AUTO-INSTALLER SETTINGS
:: # RightClickTool: true
:: # Name: FunctionInsertExtract
:: # DisplayName: Function Insert/Extract
:: # WorksOnFolders: false
:: # Background: false
:: # AllFileTypes: false
:: # Description: Extracts or inserts Excel Namespace formulas as json files
:: # Extensions: .xls, .xlsx, .xlsm
:: # AltKeyLocation: HKEY_CLASSES_ROOT\Excel.Sheet.12\shell
#>

#===============================================================================================
#|                                   POWERSHELL CODE BEGINS                                   |
#===============================================================================================

# Import required assemblies for Windows Forms GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#===============================================================================================
#|                              SCRIPT VARIABLES AND SETUP                                   |
#===============================================================================================

# Script-level variables for file paths (persistent across GUI events)
$Script:excelPath = $null    # Path to source/target Excel workbook
$Script:jsonPath = $null     # Path to JSON function definition file

# Initialize paths if launched from context menu (right-click on Excel file)
if($sourceDir) {
    $Script:excelPath = $sourceDir
    # Auto-generate JSON filename: "workbook.xlsx" → "workbook - functions.json"
    $Script:jsonPath = $sourceDir + " - functions.json"
    Write-Host "Context menu launch detected - Excel: $Script:excelPath" -ForegroundColor Green
    Write-Host "Auto-generated JSON path: $Script:jsonPath" -ForegroundColor Cyan
}
#===============================================================================================
#|                              MAIN GUI FORM CREATION                                       |
#===============================================================================================

# Create main application window
$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Function Sync - LAMBDA Import/Export Tool"
$form.Size = New-Object System.Drawing.Size(600,300)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"        # Prevent resizing for consistent layout
$form.MaximizeBox = $false                   # Disable maximize button
$form.MinimizeBox = $false                   # Disable minimize button
$form.Icon = [System.Drawing.SystemIcons]::Application

#===============================================================================================
#|                            FILE SELECTION CONTROLS                                        |
#===============================================================================================

# Excel file selection button
$btnExcel = New-Object System.Windows.Forms.Button
$btnExcel.Text = "Select Excel File"
$btnExcel.Location = New-Object System.Drawing.Point(10,20)
$btnExcel.Size = New-Object System.Drawing.Size(120,25)
$btnExcel.BackColor = [System.Drawing.Color]::LightBlue
$form.Controls.Add($btnExcel)

# Excel file path display label
$lblExcel = New-Object System.Windows.Forms.Label
$lblExcel.Text = "No Excel file selected"
$lblExcel.AutoSize = $true
$lblExcel.Location = New-Object System.Drawing.Point(150,25)
$lblExcel.ForeColor = [System.Drawing.Color]::DarkBlue
$lblExcel.MaximumSize = New-Object System.Drawing.Size(400,25)

# Update label if launched from context menu
if($sourceDir) {
    $lblExcel.Text = $sourceDir
    $lblExcel.ForeColor = [System.Drawing.Color]::Green  # Green indicates auto-detected path
}
$form.Controls.Add($lblExcel)

# JSON file selection button
$btnJson = New-Object System.Windows.Forms.Button
$btnJson.Text = "Select JSON File"
$btnJson.Location = New-Object System.Drawing.Point(10,60)
$btnJson.Size = New-Object System.Drawing.Size(120,25)
$btnJson.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnJson)

# JSON file path display label
$lblJson = New-Object System.Windows.Forms.Label
$lblJson.Text = "Default will be auto-generated"
$lblJson.AutoSize = $true
$lblJson.Location = New-Object System.Drawing.Point(150,65)
$lblJson.ForeColor = [System.Drawing.Color]::DarkGreen
$lblJson.MaximumSize = New-Object System.Drawing.Size(400,25)

# Update JSON label if launched from context menu
if($sourceDir) {
    $lblJson.Text = $sourceDir + " - functions.json"
    $lblJson.ForeColor = [System.Drawing.Color]::Green  # Green indicates auto-generated path
}
$form.Controls.Add($lblJson)

#===============================================================================================
#|                              OPERATION MODE CONTROL                                       |
#===============================================================================================

# Mode toggle checkbox - determines extract vs insert operation
$toggleMode = New-Object System.Windows.Forms.CheckBox
$toggleMode.Text = "🔄 Insert Mode (unchecked = Extract Mode)"
$toggleMode.Location = New-Object System.Drawing.Point(10,100)
$toggleMode.Size = New-Object System.Drawing.Size(300,25)
$toggleMode.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($toggleMode)

#===============================================================================================
#|                          PROGRESS TRACKING AND STATUS DISPLAY                             |
#===============================================================================================

# Progress bar for long-running operations (scanning/inserting many functions)
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,140)
$progressBar.Size = New-Object System.Drawing.Size(560,25)
$progressBar.Style = "Continuous"             # Smooth progress indication
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$form.Controls.Add($progressBar)

# Status label for operation feedback
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready - Select files and operation mode"
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(10,170)
$lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($lblStatus)

#===============================================================================================
#|                              OPERATION EXECUTION CONTROL                                  |
#===============================================================================================

# Main execution button - triggers extract or insert based on mode
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "🚀 Run Operation"
$btnRun.Location = New-Object System.Drawing.Point(10,210)
$btnRun.Size = New-Object System.Drawing.Size(120,35)
$btnRun.BackColor = [System.Drawing.Color]::LightCoral
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

#===============================================================================================
#|                            FILE DIALOG CONFIGURATION                                      |
#===============================================================================================

# Excel file browser dialog - supports all Excel formats with LAMBDA function capability
$ofdExcel = New-Object System.Windows.Forms.OpenFileDialog
$ofdExcel.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Excel Workbook (*.xlsx)|*.xlsx|Excel Macro-Enabled (*.xlsm)|*.xlsm|Excel 97-2003 (*.xls)|*.xls"
$ofdExcel.Title = "Select Excel Workbook with LAMBDA Functions"
$ofdExcel.CheckFileExists = $true
$ofdExcel.CheckPathExists = $true

# JSON file browser dialog - for function definition files
$ofdJson = New-Object System.Windows.Forms.OpenFileDialog
$ofdJson.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
$ofdJson.Title = "Select JSON Function Definition File"
$ofdJson.CheckFileExists = $true
$ofdJson.CheckPathExists = $true


#===============================================================================================
#|                              GUI EVENT HANDLERS                                           |
#===============================================================================================

# Excel file selection button click handler
$btnExcel.Add_Click({
    Write-Host "Opening Excel file selection dialog..." -ForegroundColor Yellow
    
    if($ofdExcel.ShowDialog() -eq "OK") {
        $Script:excelPath = $ofdExcel.FileName
        $lblExcel.Text = $Script:excelPath
        $lblExcel.ForeColor = [System.Drawing.Color]::Green
        
        Write-Host "Excel file selected: $Script:excelPath" -ForegroundColor Green
        
        # Auto-generate corresponding JSON path based on Excel filename
        # Example: "MyWorkbook.xlsx" → "MyWorkbook - functions.json"
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Script:excelPath)
        $folder = Split-Path $Script:excelPath -Parent
        $Script:jsonPath = Join-Path $folder "$baseName - functions.json"
        $lblJson.Text = $Script:jsonPath
        $lblJson.ForeColor = [System.Drawing.Color]::Blue  # Blue indicates auto-generated
        
        Write-Host "Auto-generated JSON path: $Script:jsonPath" -ForegroundColor Cyan
        $lblStatus.Text = "Ready - Excel file selected, JSON path auto-generated"
    }
})

# JSON file selection button click handler
$btnJson.Add_Click({
    Write-Host "Opening JSON file selection dialog..." -ForegroundColor Yellow
    
    if($ofdJson.ShowDialog() -eq "OK") {
        $Script:jsonPath = $ofdJson.FileName
        $lblJson.Text = $Script:jsonPath
        $lblJson.ForeColor = [System.Drawing.Color]::Green
        
        Write-Host "JSON file selected: $Script:jsonPath" -ForegroundColor Green
        $lblStatus.Text = "Ready - JSON file manually selected"
    }
})

#===============================================================================================
#|                              UTILITY FUNCTIONS                                            |
#===============================================================================================

<#
.SYNOPSIS
Updates progress bar and status label with visual feedback during operations.

.DESCRIPTION
Provides real-time feedback during long-running Excel COM operations. Updates both
the progress bar percentage and status text, then processes Windows messages to 
ensure GUI remains responsive during intensive operations.

.PARAMETER progress
Progress percentage (0-100) for the progress bar.

.PARAMETER text
Status message to display to user.
#>
function Update-Status {
    param(
        [Parameter(Mandatory=$true)]
        [int]$progress,
        
        [Parameter(Mandatory=$true)]
        [string]$text
    )
    
    # Update progress bar (clamp to valid range)
    $progressBar.Value = [Math]::Max(0, [Math]::Min(100, $progress))
    
    # Update status text with timestamp for operation tracking
    $timestamp = Get-Date -Format "HH:mm:ss"
    $lblStatus.Text = "[$timestamp] $text"
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    
    # Process Windows messages to keep GUI responsive during COM operations
    [System.Windows.Forms.Application]::DoEvents()
    
    # Console logging for debugging
    Write-Host "Progress: $progress% - $text" -ForegroundColor Gray
}

#===============================================================================================
#|                         MAIN OPERATION EXECUTION HANDLER                                  |
#===============================================================================================

# Main run button click handler - orchestrates extract or insert operations
$btnRun.Add_Click({
    Write-Host "" 
    Write-Host "=== Excel Function Sync Operation Started ===" -ForegroundColor Green
    
    #---------------------------------------------------------------------------
    # Input Validation
    #---------------------------------------------------------------------------
    
    # Ensure Excel file is selected
    if(-not $Script:excelPath) {
        Write-Host "ERROR: No Excel file selected" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show(
            "Please select an Excel file before running the operation.", 
            "Excel File Required", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return 
    }
    
    # Ensure JSON path is configured
    if(-not $Script:jsonPath) {
        Write-Host "ERROR: No JSON path configured" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show(
            "Please select or confirm a JSON file path.", 
            "JSON Path Required", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return 
    }
    
    # Verify Excel file exists
    if(-not (Test-Path $Script:excelPath)) {
        Write-Host "ERROR: Excel file not found: $Script:excelPath" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show(
            "The selected Excel file does not exist:`n$Script:excelPath", 
            "File Not Found", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    #---------------------------------------------------------------------------
    # Excel COM Object Initialization
    #---------------------------------------------------------------------------
    
    $modeInsert = $toggleMode.Checked
    $modeText = if($modeInsert) { "INSERT" } else { "EXTRACT" }
    
    Write-Host "Operation Mode: $modeText" -ForegroundColor Cyan
    Write-Host "Excel File: $Script:excelPath" -ForegroundColor White
    Write-Host "JSON File: $Script:jsonPath" -ForegroundColor White
    
    Update-Status 5 "Initializing Excel COM object..."
    
    try {
        # Create Excel COM object (may take a moment on first launch)
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false          # Hidden operation for performance
        $excel.DisplayAlerts = $false    # Suppress Excel dialog boxes
        $excel.ScreenUpdating = $false   # Disable screen updates for speed
        
        Write-Host "Excel COM object created successfully" -ForegroundColor Green
        Update-Status 10 "Opening workbook..."
        
        # Open the target Excel workbook
        $wb = $excel.Workbooks.Open($Script:excelPath)
        Write-Host "Workbook opened: $($wb.Name)" -ForegroundColor Green
        
    } catch {
        Write-Host "ERROR: Failed to initialize Excel COM object: $($_.Exception.Message)" -ForegroundColor Red
        Update-Status 0 "ERROR: Failed to open Excel"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to initialize Excel COM object:`n$($_.Exception.Message)", 
            "Excel COM Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    #---------------------------------------------------------------------------
    # EXTRACT MODE: Excel → JSON
    #---------------------------------------------------------------------------
    
    if(-not $modeInsert) {
        Write-Host "--- EXTRACT MODE: Scanning Excel workbook for LAMBDA functions ---" -ForegroundColor Yellow
        
        $dict = @{}                      # Dictionary to store function name → formula pairs
        $names = $wb.Names               # Get all Named Ranges in the workbook
        $count = $names.Count
        $foundFunctions = 0
        
        Write-Host "Scanning $count Named Ranges for LAMBDA functions..." -ForegroundColor Cyan
        Update-Status 15 "Scanning Named Ranges for LAMBDA functions..."
        
        # Iterate through all Named Ranges looking for LAMBDA functions
        for($i = 1; $i -le $count; $i++) {
            try {
                $nm = $names.Item($i)        # Get Named Range object
                $ref = $nm.RefersTo          # Get the formula reference
                
                # Check if this is a LAMBDA function (starts with "=LAMBDA")
                if($ref -match "^=LAMBDA") {
                    $dict[$nm.Name] = $ref   # Store function name and formula
                    $foundFunctions++
                    Write-Host "  Found LAMBDA function: $($nm.Name)" -ForegroundColor Green
                }
            } catch {
                # Some Named Ranges might be inaccessible (charts, etc.) - skip them
                Write-Host "  Skipped inaccessible Named Range at index $i" -ForegroundColor Gray
            }
            
            # Update progress (90% for scanning, 10% reserved for file operations)
            $progressPercent = [int](15 + (($i / $count) * 75))
            Update-Status $progressPercent "Scanning $i of $count Named Ranges..."
        }
        
        Write-Host "Scan complete - Found $foundFunctions LAMBDA functions" -ForegroundColor Green
        Update-Status 90 "Writing functions to JSON file..."
        
        # Export dictionary to JSON file with proper formatting
        try {
            ($dict | ConvertTo-Json -Depth 100) | Out-File -FilePath $Script:jsonPath -Encoding UTF8
            Write-Host "Successfully exported to: $Script:jsonPath" -ForegroundColor Green
            Update-Status 100 "✅ Exported $foundFunctions functions to JSON"
            
            # Show completion message
            [System.Windows.Forms.MessageBox]::Show(
                "Export completed successfully!`n`nFound: $foundFunctions LAMBDA functions`nSaved to: $Script:jsonPath", 
                "Export Complete", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } catch {
            Write-Host "ERROR: Failed to write JSON file: $($_.Exception.Message)" -ForegroundColor Red
            Update-Status 100 "❌ Failed to write JSON file"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to write JSON file:`n$($_.Exception.Message)", 
                "Export Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    #---------------------------------------------------------------------------
    # INSERT MODE: JSON → Excel
    #---------------------------------------------------------------------------
    
    else {
        Write-Host "--- INSERT MODE: Loading LAMBDA functions from JSON into Excel ---" -ForegroundColor Yellow
        
        # Verify JSON file exists
        if(Test-Path $Script:jsonPath) {
            Update-Status 15 "Reading JSON function definitions..."
            
            try {
                # Load and parse JSON file
                $jsonContent = Get-Content $Script:jsonPath -Raw
                $json = $jsonContent | ConvertFrom-Json
                $keys = $json.PSObject.Properties.Name
                $count = $keys.Count
                $insertedFunctions = 0
                $i = 0
                
                Write-Host "Loading $count functions from JSON file..." -ForegroundColor Cyan
                Write-Host "JSON file: $Script:jsonPath" -ForegroundColor White
                
                # Iterate through each function definition in JSON
                foreach($functionName in $keys) {
                    $i++
                    $formula = $json.$functionName
                    
                    try {
                        # Delete existing function with same name (prevents duplicates)
                        try { 
                            $wb.Names.Item($functionName).Delete() 
                            Write-Host "  Replaced existing function: $functionName" -ForegroundColor Yellow
                        } catch { 
                            # Function doesn't exist yet - this is normal for new functions
                        }
                        
                        # Add new Named Range with LAMBDA formula
                        $wb.Names.Add($functionName, $formula) | Out-Null
                        $insertedFunctions++
                        Write-Host "  Inserted function: $functionName" -ForegroundColor Green
                        
                    } catch {
                        Write-Host "  ERROR inserting function '$functionName': $($_.Exception.Message)" -ForegroundColor Red
                    }
                    
                    # Update progress (75% for insertion, 15% reserved for saving)
                    $progressPercent = [int](15 + (($i / $count) * 75))
                    Update-Status $progressPercent "Inserted $i of $count functions..."
                }
                
                Write-Host "Function insertion complete - Saving workbook..." -ForegroundColor Cyan
                Update-Status 90 "Saving workbook with new functions..."
                
                # Save the workbook with new functions
                $wb.Save()
                Write-Host "Workbook saved successfully" -ForegroundColor Green
                Update-Status 100 "✅ Inserted $insertedFunctions functions into Excel"
                
                # Show completion message
                [System.Windows.Forms.MessageBox]::Show(
                    "Insert completed successfully!`n`nInserted: $insertedFunctions LAMBDA functions`nSaved to: $Script:excelPath", 
                    "Insert Complete", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                
            } catch {
                Write-Host "ERROR: Failed to process JSON file: $($_.Exception.Message)" -ForegroundColor Red
                Update-Status 100 "❌ Failed to process JSON file"
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to process JSON file:`n$($_.Exception.Message)", 
                    "Insert Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            
        } else {
            Write-Host "ERROR: JSON file not found: $Script:jsonPath" -ForegroundColor Red
            Update-Status 100 "❌ JSON file not found"
            [System.Windows.Forms.MessageBox]::Show(
                "JSON file not found:`n$Script:jsonPath`n`nPlease check the file path or run Extract mode first.", 
                "File Not Found", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    }

    #---------------------------------------------------------------------------
    # Excel COM Object Cleanup (Critical for preventing Excel process leaks)
    #---------------------------------------------------------------------------
    
    Write-Host "Cleaning up Excel COM objects..." -ForegroundColor Yellow
    
    try {
        # Close workbook (save changes if in insert mode)
        $wb.Close($true)
        Write-Host "Workbook closed" -ForegroundColor Green
        
        # Quit Excel application
        $excel.Quit()
        Write-Host "Excel application closed" -ForegroundColor Green
        
        # Release COM objects to prevent memory leaks and lingering Excel processes
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        # Force garbage collection to clean up COM references
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        [gc]::Collect()  # Second collection recommended for COM objects
        
        Write-Host "COM object cleanup completed" -ForegroundColor Green
        
    } catch {
        Write-Host "WARNING: Error during COM cleanup: $($_.Exception.Message)" -ForegroundColor Yellow
        # Continue execution - cleanup errors are usually not critical
    }
    
    Write-Host "=== Excel Function Sync Operation Completed ===" -ForegroundColor Green
    Write-Host ""
})

#===============================================================================================
#|                              APPLICATION STARTUP                                          |
#===============================================================================================

Write-Host "Starting Excel Function Sync Tool..." -ForegroundColor Green
Write-Host "Purpose: Extract/Insert LAMBDA functions between Excel and JSON" -ForegroundColor Cyan

# Show the main form as modal dialog
[void]$form.ShowDialog()

Write-Host "Excel Function Sync Tool closed" -ForegroundColor Yellow