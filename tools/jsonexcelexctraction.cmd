# 2>nul & @echo off & PowerShell -ExecutionPolicy Bypass -WindowStyle hidden -Command "$s = Get-Content '%~dpnx0' | Select-Object -Skip 1 | Out-String; $s='Set-Location ""%~dp0"";'+$s; $sourceDir = '%~1'; Invoke-Expression $s" & goto :eof

<#
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Vars

$Script:excelPath = $null
$Script:jsonPath = $null
if($sourceDir){
    $Script:excelPath = $sourceDir
      $Script:jsonPath = $sourceDir + " - functions.json"
}
$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Function Sync"
$form.Size = New-Object System.Drawing.Size(600,300)
$form.StartPosition = "CenterScreen"

# File selectors
$btnExcel = New-Object System.Windows.Forms.Button
$btnExcel.Text = "Select Excel File"
$btnExcel.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($btnExcel)

$lblExcel = New-Object System.Windows.Forms.Label
$lblExcel.Text = "No Excel file selected"
$lblExcel.AutoSize = $true
$lblExcel.Location = New-Object System.Drawing.Point(150,25)
if($sourceDir){
      $lblExcel.Text = $sourceDir
}
$form.Controls.Add($lblExcel)

$btnJson = New-Object System.Windows.Forms.Button
$btnJson.Text = "Select JSON File"
$btnJson.Location = New-Object System.Drawing.Point(10,60)
$form.Controls.Add($btnJson)

$lblJson = New-Object System.Windows.Forms.Label
$lblJson.Text = "Default will be auto-generated"
if($sourceDir){
      $lblJson.Text = $sourceDir + "- functions.json"
}
$lblJson.AutoSize = $true
$lblJson.Location = New-Object System.Drawing.Point(150,65)
$form.Controls.Add($lblJson)

# Mode toggle
$toggleMode = New-Object System.Windows.Forms.CheckBox
$toggleMode.Text = "Insert Mode (unchecked = Extract Mode)"
$toggleMode.Location = New-Object System.Drawing.Point(10,100)
$form.Controls.Add($toggleMode)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,140)
$progressBar.Size = New-Object System.Drawing.Size(560,25)
$form.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Idle"
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(10,170)
$form.Controls.Add($lblStatus)

# Run button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(10,210)
$form.Controls.Add($btnRun)

# File dialogs
$ofdExcel = New-Object System.Windows.Forms.OpenFileDialog
$ofdExcel.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"

$ofdJson = New-Object System.Windows.Forms.OpenFileDialog
$ofdJson.Filter = "JSON Files (*.json)|*.json"


$btnExcel.Add_Click({
    if($ofdExcel.ShowDialog() -eq "OK"){
        $Script:excelPath = $ofdExcel.FileName
        $lblExcel.Text = $Script:excelPath
            Write-Host $Script:excelPath;
        # Default JSON path
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Script:excelPath)
        $folder   = Split-Path $Script:excelPath -Parent
        $Script:jsonPath = Join-Path $folder "$baseName - functions.json"
        $lblJson.Text = $Script:jsonPath
    }
})

$btnJson.Add_Click({
    if($ofdJson.ShowDialog() -eq "OK"){
        $Script:jsonPath = $ofdJson.FileName
        $lblJson.Text = $Script:jsonPath
    }
})

function Update-Status {
    param($progress,$text)
    $progressBar.Value = $progress
    $lblStatus.Text = $text
    [System.Windows.Forms.Application]::DoEvents()
}

$btnRun.Add_Click({
    Write-Host "Run button clicked."
    if (-not $Script:excelPath) {
        Write-Host "No Excel file selected."
        [System.Windows.Forms.MessageBox]::Show("Please select an Excel file.")
        return 
    }
    if (-not (Test-Path $Script:excelPath)) {
        Write-Host "Excel file does not exist: $Script:excelPath"
        Update-Status 0 "Excel file not found."
        return
    }

    if (-not $Script:jsonPath) {
        Write-Host "No JSON file path set."
        [System.Windows.Forms.MessageBox]::Show("Please select or confirm a JSON path.")
        return 
    }

    Write-Host "Opening Excel application..."
    Update-Status 5 "Launching Excel..."
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
    } catch {
        Write-Host ("Failed to open Excel COM object: " + $_)
        Update-Status 0 "Failed to open Excel."
        return
    }

    Write-Host "Opening workbook: $Script:excelPath"
    Update-Status 10 "Opening workbook..."
    try {
        $wb = $excel.Workbooks.Open($Script:excelPath)
    } catch {
        Write-Host ("Failed to open workbook: " + $_)
        Update-Status 0 "Workbook open failed."
        $excel.Quit()
        return
    }

    $modeInsert = $toggleMode.Checked
    Write-Host "Mode: " + ($(if ($modeInsert) { "Insert" } else { "Extract" }))

    if (-not $modeInsert) {
        # Extract Mode
        Write-Host "Extracting named formulas..."
        Update-Status 20 "Extracting formulas..."
        $dict = @{}
        $names = $wb.Names
        $count = $names.Count
        Write-Host "Found $count named items."
		Write-Host "Workbook-Level Names Count: $($wb.Names.Count)"
foreach ($n in $wb.Names) {
    try {
        Write-Host "WB Name: $($n.Name) => $($n.RefersTo)"
    } catch {
        Write-Host ("[Workbook] Error reading name: " + $_)
    }
}

        for ($i = 1; $i -le $count; $i++) {
            try {
                $nm = $names.Item($i)
                $ref = $nm.RefersTo
                if ($ref -match "^= ?LAMBDA") {
                    Write-Host "Found LAMBDA: $($nm.Name)"
                    $dict[$nm.Name] = $ref
                }
            } catch {
                Write-Host ("Error reading name item " + $i + ":" + $_)
            }
            Update-Status ([int](($i / $count) * 90)) "Scanning $i of $count..."
        }

        try {
            ($dict | ConvertTo-Json -Depth 100) | Out-File -FilePath $Script:jsonPath -Encoding UTF8
            Write-Host "Exported JSON to: $Script:jsonPath"
            Update-Status 100 "Exported to $Script:jsonPath"
        } catch {
            Write-Host ("Failed to save JSON: " + $_)
            Update-Status 0 "Failed to write JSON."
        }

    } else {
        # Insert Mode
        Write-Host "Inserting functions from JSON: $Script:jsonPath"
        Update-Status 20 "Reading JSON file..."

        if (Test-Path $Script:jsonPath) {
            try {
                $json = Get-Content $Script:jsonPath -Raw | ConvertFrom-Json
            } catch {
                Write-Host ("Failed to read/parse JSON: " + $_)
                Update-Status 0 "JSON read failed."
                return
            }

            $keys = $json.PSObject.Properties.Name
            $count = $keys.Count
            Write-Host "Found $count entries in JSON."

            $i = 0
            foreach ($k in $keys) {
                $i++
                $formula = $json.$k
                Write-Host "Inserting $k = $formula"
                try { $wb.Names.Item($k).Delete() } catch {}
                try {
                    $wb.Names.Add($k, $formula) | Out-Null
                } catch {
                    Write-Host ("Error inserting " + $k + ":" + $_)
                }
                Update-Status ([int](($i / $count) * 90)) "Inserted $i of $count..."
            }

            try {
                $wb.Save()
                Write-Host "Workbook saved: $Script:excelPath"
                Update-Status 100 "Inserted functions into $Script:excelPath"
            } catch {
                Write-Host ("Error saving workbook: " + $_)
                Update-Status 0 "Workbook save failed."
            }

        } else {
            Write-Host "JSON file does not exist: $Script:jsonPath"
            Update-Status 100 "JSON file not found."
        }
    }

    Write-Host "Cleaning up..."
    $wb.Close($true)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)| Out-Null
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
    Write-Host "Finished."
})



[void]$form.ShowDialog()
