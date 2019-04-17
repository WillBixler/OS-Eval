# Import Scan

$Desktop = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$Path = "$($Desktop)\OSUpgradeScan.csv"
$computers = Import-Csv -Path $Path

# Create Excel Sheet

$excel = New-Object -ComObject excel.application
$excel.visible = $true

$workbook = $excel.Workbooks.Add()
$UpgradeEvalutation = $workbook.Worksheets.Item(1)
$UpgradeEvalutation.Name = "Upgrade Evaluation"

# Header

$UpgradeEvalutation.Cells.Item(1,1) = "Computer Name"
$UpgradeEvalutation.Cells.Item(1,2) = "Current OS"
$UpgradeEvalutation.Cells.Item(1,3) = "OS Architecture"
$UpgradeEvalutation.Cells.Item(1,4) = "Processor"
$UpgradeEvalutation.Cells.Item(1,5) = "Ram"

# Populate Data

$row = 2
foreach ($computer in $computers) {
    $UpgradeEvalutation.Cells.Item($row, 1) = $computer.Name
    if ($computer.OSCurrent -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 4
    } else {
        if ($computer.ArchitectureValid -eq "TRUE" -and $computer.ProcessorValid -eq "TRUE" -and $computer.RamValid -eq "TRUE") {
            $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 6
        } else {
            $UpgradeEvalutation.Cells.Item($row,1).Interior.ColorIndex = 3
        }
    }
    $UpgradeEvalutation.Cells.Item($row, 2) = $computer.OSName
    if ($computer.OSCurrent -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,2).Interior.ColorIndex = 4
    } else {
        $UpgradeEvalutation.Cells.Item($row,2).Interior.ColorIndex = 6
    }
    $UpgradeEvalutation.Cells.Item($row, 3) = $computer.Architecture
    if ($computer.ArchitectureValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,3).Interior.ColorIndex = 4
    } else {
        $UpgradeEvalutation.Cells.Item($row,3).Interior.ColorIndex = 3
    }
    $UpgradeEvalutation.Cells.Item($row, 4) = $computer.Processor
    if ($computer.ProcessorValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,4).Interior.ColorIndex = 4
    } else {
        $UpgradeEvalutation.Cells.Item($row,4).Interior.ColorIndex = 3
    }
    $UpgradeEvalutation.Cells.Item($row, 5) = "$($computer.Ram)GB"
    if ($computer.RamValid -eq "TRUE") {
        $UpgradeEvalutation.Cells.Item($row,5).Interior.ColorIndex = 4
    } else {
        $UpgradeEvalutation.Cells.Item($row,5).Interior.ColorIndex = 3
    }
    $row++
}