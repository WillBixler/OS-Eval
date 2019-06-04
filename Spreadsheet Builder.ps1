# Imports

Import-Module "$PSScriptRoot\Validator.psm1" -Force

# Import Scan

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "CSV (*.csv)|*.csv"
}
$FileBrowser.ShowDialog()

$Path = $FileBrowser.FileName
$FileName = $FileBrowser.SafeFileName.Split(".")[0]
Write-Host "File: $($FileName)" -ForegroundColor Green
$computers = Import-Csv -Path $Path

# Create Excel Sheet

Write-Host "Creating Spreadsheet..." -ForegroundColor Green

$excel = New-Object -ComObject excel.application
$excel.visible = $true

$workbook = $excel.Workbooks.Add()
$UpgradeEvalutation = $workbook.Worksheets.Item(1)
$UpgradeEvalutation.Name = "Upgrade Evaluation"

# Header

$UpgradeEvalutation.Cells[1,1].EntireRow.Font.Size = 13
$UpgradeEvalutation.Cells[1,1].EntireRow.Font.Bold = $true
$UpgradeEvalutation.Cells.Item(1,1) = "Computer Name"
$UpgradeEvalutation.Cells.Item(1,2) = "Current OS"
$UpgradeEvalutation.Cells.Item(1,3) = "OS Architecture"
$UpgradeEvalutation.Cells.Item(1,4) = "Processor"
$UpgradeEvalutation.Cells.Item(1,5) = "Ram"
$UpgradeEvalutation.Cells.Item(1,6) = "Age"
$UpgradeEvalutation.Cells.Item(1,7) = "Hard Drive Capacity"
$UpgradeEvalutation.Cells.Item(1,8) = "SSD"
$UpgradeEvalutation.Cells.Item(1,9) = "Suggested Action"
$UpgradeEvalutation.Cells.Item(1,11) = "Error"


# Populate Data
$OKCount = 0
$OSUpgradeCount = 0
$PartsAndOSUpgradeCount = 0
$ReplacePartsCount = 0
$ReplaceComputerCount = 0

$row = 2
foreach ($computer in $computers) {

    $UpgradeEvalutation.Cells.Item($row, 1) = $computer.Name
    $UpgradeEvalutation.Cells.Item($row, 2) = $computer.OSName
    $UpgradeEvalutation.Cells.Item($row, 3) = $computer.Architecture
    $UpgradeEvalutation.Cells.Item($row, 4) = $computer.Processor
    if ($computer.Ram) {
        $UpgradeEvalutation.Cells.Item($row, 5) = "$($computer.Ram)GB"
    }
    if ($computer.Age) {
        $UpgradeEvalutation.Cells.Item($row, 6) = "$($computer.Age) years"
    }
    if ($computer.HardDriveSize) {
        $UpgradeEvalutation.Cells.Item($row, 7) = "$($computer.HardDriveSize)GB"
    }
    $UpgradeEvalutation.Cells.Item($row, 8) = $computer.HardDriveModel
    if ($computer.Error) {
        $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 15
        $UpgradeEvalutation.Cells.Item($row, 11) = "$($computer.Error)"
        $UpgradeEvalutation.Cells.Item($row,11).Interior.ColorIndex = 3
        $UpgradeEvalutation.Cells.Item($row,11).Font.ColorIndex = 2
        $UpgradeEvalutation.Cells.Item($row,11).Font.Bold = $true
    }
    if ($computer.Error -like "*Offline*" -or $computer.Error -like "*RPC server unavailable*") {
        $UpgradeEvalutation.Cells.EntireRow[$row].Interior.ColorIndex = 15
        $UpgradeEvalutation.Cells.EntireRow[$row].Font.Bold = $true
    } else {

        $computer = Validate-Computer -computer $computer
        
        if ($computer.OSCurrent) {
            $UpgradeEvalutation.Cells.Item($row,2).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row,2).Interior.ColorIndex = 6
        }
        
        if ($computer.ArchitectureValid) {
            $UpgradeEvalutation.Cells.Item($row,3).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row,3).Interior.ColorIndex = 3
            $UpgradeEvalutation.Cells.Item($row,3).Font.ColorIndex = 2
            $UpgradeEvalutation.Cells.Item($row,3).Font.Bold = $true
        }
        
        if ($computer.ProcessorValid) {
            $UpgradeEvalutation.Cells.Item($row,4).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row,4).Interior.ColorIndex = 3
            $UpgradeEvalutation.Cells.Item($row,4).Font.ColorIndex = 2
            $UpgradeEvalutation.Cells.Item($row,4).Font.Bold = $true
        }
        
        if ($computer.RamValid) {
            $UpgradeEvalutation.Cells.Item($row,5).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row,5).Interior.ColorIndex = 6
        }
        
        if ($computer.AgeValid) {
            $UpgradeEvalutation.Cells.Item($row,6).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row,6).Interior.ColorIndex = 3
            $UpgradeEvalutation.Cells.Item($row,6).Font.ColorIndex = 2
            $UpgradeEvalutation.Cells.Item($row,6).Font.Bold = $true
        }
        
        if ($computer.HardDriveValid) {
            $UpgradeEvalutation.Cells.Item($row,7).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row,7).Interior.ColorIndex = 6
        }
        
        if ($computer.SSD) {
            $UpgradeEvalutation.Cells.Item($row,8).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row,8).Interior.ColorIndex = 6
        }

        switch ($computer.SuggestedAction) {
            -1 { 
                $ReplaceComputerCount++
                $UpgradeEvalutation.Cells.Item($row, 9) = "Replace Computer"
                $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 3
                $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 3
                $UpgradeEvalutation.Cells.Item($row,1).Font.ColorIndex = 2
                $UpgradeEvalutation.Cells.Item($row,9).Font.ColorIndex = 2
                $UpgradeEvalutation.Cells.Item($row,1).Font.Bold = $true
                $UpgradeEvalutation.Cells.Item($row,9).Font.Bold = $true
            }
            0 {
                $OKCount++
                $UpgradeEvalutation.Cells.Item($row, 9) = "OK"
                $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 4
                $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 4
            }
            1 {
                $OSUpgradeCount++
                $UpgradeEvalutation.Cells.Item($row, 9) = "Upgrade OS"
                $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 6
                $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 6
            }
            2 {
                $ReplacePartsCount++
                $UpgradeEvalutation.Cells.Item($row, 9) = "Replace Parts"
                $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 6
                $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 6
            }
            3 {
                $PartsAndOSUpgradeCount++
                $UpgradeEvalutation.Cells.Item($row, 9) = "Upgrade OS & Replace Parts"
                $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 6
                $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 6
            }
        }
    }
    
    $row++
}

$row += 2
$UpgradeEvalutation.Cells.Item($row, 1) = "Totals"
$UpgradeEvalutation.Cells.Item($row, 1).Font.Size = 14
$UpgradeEvalutation.Cells.Item($row, 1).Font.Bold = $true
$row++
$UpgradeEvalutation.Cells.Item($row, 1) = "OK"
$UpgradeEvalutation.Cells.Item($row, 2) = $OKCount
$row++
$UpgradeEvalutation.Cells.Item($row, 1) = "Upgrade OS"
$UpgradeEvalutation.Cells.Item($row, 2) = $OSUpgradeCount
$row++
$UpgradeEvalutation.Cells.Item($row, 1) = "Replace Parts"
$UpgradeEvalutation.Cells.Item($row, 2) = $ReplacePartsCount
$row++
$UpgradeEvalutation.Cells.Item($row, 1) = "Upgrade OS & Replace Parts"
$UpgradeEvalutation.Cells.Item($row, 2) = $PartsAndOSUpgradeCount
$row++
$UpgradeEvalutation.Cells.Item($row, 1) = "Replace Computer"
$UpgradeEvalutation.Cells.Item($row, 2) = $ReplaceComputerCount

# Auto fit cells
$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit()

Write-Host "Formating complete... Please select where to save the file..."

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    FileName = $FileName
}
$cancel = $FileBrowser.ShowDialog() -eq "cancel"

if (!$cancel) {
    $Path = $FileBrowser.FileName
    $UpgradeEvalutation.SaveAs($Path)
}