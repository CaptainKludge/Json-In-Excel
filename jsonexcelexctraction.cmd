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
$ofdExcel.Filter = "Excel Files (*.xlsx)|*.xlsx"

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
    if(-not $Script:excelPath){
        Write-Host $Script:excelPath;
        [System.Windows.Forms.MessageBox]::Show("Please select an Excel file.");
        return 
    }
    if(-not $Script:jsonPath){ [System.Windows.Forms.MessageBox]::Show("Please select or confirm a JSON path."); return }

    $modeInsert = $toggleMode.Checked
    Update-Status 0 "Opening Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($Script:excelPath)

    if(-not $modeInsert){
        # Extract Mode
        $dict = @{}
        $names = $wb.Names
        $count = $names.Count
        for($i=1;$i -le $count;$i++){
            try {
                $nm = $names.Item($i)
                $ref = $nm.RefersTo
                if($ref -match "^=LAMBDA"){
                    $dict[$nm.Name] = $ref
                }
            } catch {}
            Update-Status ([int](($i/$count)*90)) "Scanning $i of $count..."
        }
        ($dict | ConvertTo-Json -Depth 100) | Out-File -FilePath $Script:jsonPath -Encoding UTF8
        Update-Status 100 "Exported to $Script:jsonPath"
    }
    else {
        # Insert Mode
        if(Test-Path $Script:jsonPath){
            $json = Get-Content $Script:jsonPath -Raw | ConvertFrom-Json
            $keys = $json.PSObject.Properties.Name
            $count = $keys.Count
            $i=0
            foreach($k in $keys){
                $i++
                $formula = $json.$k
                try { $wb.Names.Item($k).Delete() } catch {}
                $wb.Names.Add($k,$formula) | Out-Null
                Update-Status ([int](($i/$count)*90)) "Inserted $i of $count..."
            }
            $wb.Save()
            Update-Status 100 "Inserted functions into $Script:excelPath"
        } else {
            Update-Status 100 "JSON file not found."
        }
    }

    $wb.Close($true)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)| Out-Null
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
})

[void]$form.ShowDialog()