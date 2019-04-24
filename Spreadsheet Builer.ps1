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

$row = 2
foreach ($computer in $computers) {
    $UpgradeEvalutation.Cells.Item($row, 1) = $computer.Name
    if ($computer.Error) {
        $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 15
        $UpgradeEvalutation.Cells.Item($row, 11) = "$($computer.Error)"
        $UpgradeEvalutation.Cells.Item($row,11).Interior.ColorIndex = 3
        $UpgradeEvalutation.Cells.Item($row,11).Font.ColorIndex = 2
        $UpgradeEvalutation.Cells.Item($row,11).Font.Bold = $true
    }
    $UpgradeEvalutation.Cells.Item($row, 2) = $computer.OSName
    if ($computer.OSCurrent -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,2).Interior.ColorIndex = 4
    } elseif ($computer.OSCurrent -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row,2).Interior.ColorIndex = 6
    }
    $UpgradeEvalutation.Cells.Item($row, 3) = $computer.Architecture
    if ($computer.ArchitectureValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,3).Interior.ColorIndex = 4
    } elseif ($computer.ArchitectureValid -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row,3).Interior.ColorIndex = 3
        $UpgradeEvalutation.Cells.Item($row,3).Font.ColorIndex = 2
        $UpgradeEvalutation.Cells.Item($row,3).Font.Bold = $true
    }
    $UpgradeEvalutation.Cells.Item($row, 4) = $computer.Processor
    if ($computer.ProcessorValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,4).Interior.ColorIndex = 4
    } elseif ($computer.ProcessorValid -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row,4).Interior.ColorIndex = 6
    }
    $UpgradeEvalutation.Cells.Item($row, 5) = "$($computer.Ram)GB"
    if ($computer.RamValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,5).Interior.ColorIndex = 4
    } elseif ($computer.RamValid -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row,5).Interior.ColorIndex = 6
    } else {
        $UpgradeEvalutation.Cells.Item($row, 5) = ""
    }
    $UpgradeEvalutation.Cells.Item($row, 6) = "$($computer.Age) years"
    if ($computer.AgeValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,6).Interior.ColorIndex = 4
    } elseif ($computer.AgeValid -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row,6).Interior.ColorIndex = 3
        $UpgradeEvalutation.Cells.Item($row,6).Font.ColorIndex = 2
        $UpgradeEvalutation.Cells.Item($row,6).Font.Bold = $true
    } else {
        $UpgradeEvalutation.Cells.Item($row, 6) = ""
    }
    $UpgradeEvalutation.Cells.Item($row, 7) = "$($computer.HardDriveSize)GB"
    if ($computer.HardDriveValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,7).Interior.ColorIndex = 4
    } elseif ($computer.HardDriveValid -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row,7).Interior.ColorIndex = 6
    }
    $UpgradeEvalutation.Cells.Item($row, 8) = $computer.HardDriveModel
    if ($computer.SSD -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,8).Interior.ColorIndex = 4
    } elseif ($computer.SSD -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row,8).Interior.ColorIndex = 6
    }
    if ($computer.ArchitectureValid -eq "FALSE" -or $computer.AgeValid -eq "FALSE") {
        $UpgradeEvalutation.Cells.Item($row, 9) = "Replace"
        $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 3
        $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 3
        $UpgradeEvalutation.Cells.Item($row,1).Font.ColorIndex = 2
        $UpgradeEvalutation.Cells.Item($row,9).Font.ColorIndex = 2
        $UpgradeEvalutation.Cells.Item($row,1).Font.Bold = $true
        $UpgradeEvalutation.Cells.Item($row,9).Font.Bold = $true
    } else {
        if ($computer.OSCurrent -eq "TRUE" -and $computer.ProcessorValid -eq "TRUE" -and $computer.RamValid -eq "TRUE") {
            $UpgradeEvalutation.Cells.Item($row, 9) = "OK"
            $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 4
            $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 4
        } else {
            $UpgradeEvalutation.Cells.Item($row, 9) = "Upgrade"
            $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 6
            $UpgradeEvalutation.Cells.Item($row,9).Interior.ColorIndex = 6
        }
    }
    $row++
}

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