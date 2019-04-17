# Get computers

$computers = @()

$inputUN = Read-Host -Prompt "Input the domain administrator`'s user name"
$AdminAccount = "$($env:USERDNSDOMAIN)\$($inputUN)"
$AdminAccount
$inputPW = Read-Host -Prompt "Input the domain administrator`'s password"
$password = ConvertTo-SecureString $inputPW -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential $AdminAccount, $password

Write-Host "Getting network computers... This can take a while..." -ForegroundColor Green
$networkComputers = (([adsi]"WinNT://$((Get-WMIObject Win32_ComputerSystem).Domain)").Children).Where({$_.schemaclassname -eq 'computer'})

# Data Collection

foreach ($RemoteComputer in $networkComputers) {

    $RemoteComputerName = $RemoteComputer.Path.Split("/")[3]
    Write-Host "`nScanning $($RemoteComputerName)" -ForegroundColor Green
    
    if ($ComputerSystem = Get-WmiObject Win32_ComputerSystem -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction SilentlyContinue) {

        $OS = Get-WmiObject Win32_OperatingSystem -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction SilentlyContinue
        $Processor = Get-WmiObject Win32_processor -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction SilentlyContinue

        $computer = New-Object -TypeName psobject

        $computer | Add-Member -MemberType NoteProperty -Name Name -Value $ComputerSystem.Name
        $computer | Add-Member -MemberType NoteProperty -Name OS -Value $OS.Version
        $computer | Add-Member -MemberType NoteProperty -Name OSName -Value $OS.Name.Split("|")[0]
        $computer | Add-Member -MemberType NoteProperty -Name Architecture -Value $OS.OSArchitecture
        $computer | Add-Member -MemberType NoteProperty -Name Processor -Value $Processor.Name
        $computer | Add-Member -MemberType NoteProperty -Name Ram -Value ([Math]::Round($ComputerSystem.TotalPhysicalMemory/1GB,0))

        # Begin Evaluation

        $computer | Add-Member -MemberType NoteProperty -Name OSCurrent -Value $false
        $computer | Add-Member -MemberType NoteProperty -Name ArchitectureValid -Value $false
        $computer | Add-Member -MemberType NoteProperty -Name ProcessorValid -Value $false
        $computer | Add-Member -MemberType NoteProperty -Name RamValid -Value $false

        if ($OS.Version -like "10.*") {
            $computer.OSCurrent = $true
        }

        if ($OS.OSArchitecture -like "64-bit") {
            $computer.ArchitectureValid = $true
        }

        $Processor_Gen = $computer.Processor.Substring(19,1)
        if ($Processor_Gen -ge 5) {
            $computer.ProcessorValid = $true
        }

        if ($computer.Ram -ge 8) {
            $computer.RamValid = $true
        }

        # Begin Output

        if ($computer.ArchitectureValid -and $computer.ProcessorValid -and $computer.RamValid) {
            if ($computer.OSCurrent) {
                Write-Host $computer.Name -ForegroundColor Green
            } else {
                Write-Host $computer.Name -ForegroundColor Yellow
            }
        } else {
            Write-Host $computer.Name -ForegroundColor Red
        }

        if ($computer.OSCurrent) {
            Write-Host $computer.OSName -ForegroundColor Green
        } else {
            Write-Host $computer.OSName -ForegroundColor Yellow
        }

        if ($computer.ArchitectureValid) {
            Write-Host $computer.Architecture -ForegroundColor Green
        } else {
            Write-Host $computer.Architecture -ForegroundColor Red
        }

        if ($computer.ProcessorValid) {
            Write-Host $computer.Processor -ForegroundColor Green
        } else {
            Write-Host $computer.Processor -ForegroundColor Red
        }

        if ($computer.RamValid) {
            Write-Host "$($computer.Ram)GB" -ForegroundColor Green
        } else {
            Write-Host "$($computer.Ram)GB" -ForegroundColor Red
        }

        $computers += $computer
    }

}

$Desktop = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$Path = "$($Desktop)\OSUpgradeScan.csv"
$computers | Export-Csv -Path $Path

Write-Host "`n`nCSV Exported to Desktop - $($Path)"

pause