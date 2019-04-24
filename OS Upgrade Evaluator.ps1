# Init Variables

$computers = @()

# Get network credentials

Add-Type -assembly System.Windows.Forms
$CredentialsForm = New-Object System.Windows.Forms.Form
$CredentialsForm.Text = "Network Administrator Credentials"
$CredentialsForm.Width = 320
$CredentialsForm.Height = 200
$CredentialsForm.AutoSize = $true
$CredentialsForm.StartPosition = "CenterScreen"

$instructionsLabel = New-Object System.Windows.Forms.Label
$instructionsLabel.Location = New-Object System.Drawing.Size(10,20)
$instructionsLabel.Size = New-Object System.Drawing.Size(300,30)
$instructionsLabel.Text = "Please input the network administrator`'s login credentials."
$CredentialsForm.Controls.Add($instructionsLabel)

$UnLabel = New-Object System.Windows.Forms.Label
$UnLabel.Location = New-Object System.Drawing.Size(10,50)
$UnLabel.Size = New-Object System.Drawing.Size(80,30)
$UnLabel.Text = "Username:"
$CredentialsForm.Controls.Add($UnLabel)

$PwLabel = New-Object System.Windows.Forms.Label
$PwLabel.Location = New-Object System.Drawing.Size(10,80)
$PwLabel.Size = New-Object System.Drawing.Size(80,30)
$PwLabel.Text = "Password:"
$CredentialsForm.Controls.Add($PwLabel)

$UnInput = New-Object System.Windows.Forms.TextBox
$UnInput.Location = New-Object System.Drawing.Size(100,50)
$UnInput.Size = New-Object System.Drawing.Size(200,20)
if ($env:USERDNSDOMAIN){
    $UnInput.Text = "$($env:USERDNSDOMAIN)\$($env:USERNAME)"
} else {
    $UnInput.Text = $env:USERNAME
}
$CredentialsForm.Controls.Add($UnInput)

$PwInput = New-Object System.Windows.Forms.MaskedTextBox
$PwInput.PasswordChar = "*"
$PwInput.Location = New-Object System.Drawing.Size(100,80)
$PwInput.Size = New-Object System.Drawing.Size(200,20)
$CredentialsForm.Controls.Add($PwInput)

$CurrentCredentialsButton = New-Object System.Windows.Forms.CheckBox
$CurrentCredentialsButton.Text = "Use current user`'s credentials"
$CurrentCredentialsButton.Location = New-Object System.Drawing.Size(100,100)
$CurrentCredentialsButton.Size = New-Object System.Drawing.Size(200,20)
$CurrentCredentialsButton.Add_Click({
    if ($CurrentCredentialsButton.Checked) {
        $UnInput.Enabled = $false
        $PwInput.Enabled = $false
    } else {
        $UnInput.Enabled = $true
        $PwInput.Enabled = $true
    }

})
$CredentialsForm.Controls.Add($CurrentCredentialsButton)

$OkButton = New-Object System.Windows.Forms.Button
$OkButton.Location = New-Object System.Drawing.Size(160,120)
$OkButton.Size = New-Object System.Drawing.Size(130,30)
$OkButton.Text = "OK"
$OkButton.Add_Click({
    if ($CurrentCredentialsButton.Checked) {
        $Script:UseCurrentCredentials = $true
    } else {
        $script:inputUN = $UnInput.Text
        $script:inputPW = $PwInput.Text
    }
    $CredentialsForm.Close()
})
$CredentialsForm.Controls.Add($OkButton)

[void]$CredentialsForm.ShowDialog()
if (!$UseCurrentCredentials) {
    $AdminAccount = $inputUN
    $AdminPassword = ConvertTo-SecureString $inputPW -AsPlainText -Force
    $credentials = New-Object System.Management.Automation.PSCredential $AdminAccount, $AdminPassword
}

# Get computers

Write-Host "Scanning network computers... This can take a while..." -ForegroundColor Green
$networkComputers = (([adsi]"WinNT://$((Get-WMIObject Win32_ComputerSystem).Domain)").Children).Where({$_.schemaclassname -eq 'computer'})

# Data Collection

$localComputerName = (Get-WmiObject Win32_ComputerSystem).Name

foreach ($RemoteComputer in $networkComputers) {

    $RemoteComputerName = $RemoteComputer.Path.Split("/")[3]
    Write-Host "`nScanning $($RemoteComputerName)..." -ForegroundColor Green

    $computer = New-Object -TypeName psobject

    $computer | Add-Member -MemberType NoteProperty -Name Name -Value $RemoteComputerName
    $computer | Add-Member -MemberType NoteProperty -Name OS -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name OSName -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name Architecture -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name Processor -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name Ram -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name Age -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name HardDriveSize -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name HardDriveModel -Value $null
    $computer | Add-Member -MemberType NoteProperty -Name OSCurrent -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name ArchitectureValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name ProcessorValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name RamValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name AgeValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name HardDriveValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name SSD -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name Error -Value $null
    
    try {

        if ($RemoteComputerName -eq $localComputerName) {
            $ComputerSystem = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
            $OS = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
            $Processor = Get-WmiObject Win32_processor -ErrorAction Stop
            $HardDrive = Get-WmiObject Win32_DiskDrive -ErrorAction Stop
        } elseif ($UseCurrentCredentials) {
            $ComputerSystem = Get-WmiObject Win32_ComputerSystem -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop
            $OS = Get-WmiObject Win32_OperatingSystem -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop
            $Processor = Get-WmiObject Win32_processor -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop
            $HardDrive = Get-WmiObject Win32_DiskDrive -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop
        } else {
            $ComputerSystem = Get-WmiObject Win32_ComputerSystem -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop
            $OS = Get-WmiObject Win32_OperatingSystem -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop
            $Processor = Get-WmiObject Win32_processor -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop
            $HardDrive = Get-WmiObject Win32_DiskDrive -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop
        }

        $computer.Name = $ComputerSystem.Name
        $computer.OS = $OS.Version
        $computer.OSName = $OS.Name.Split("|")[0]
        $computer.Architecture = $OS.OSArchitecture
        $computer.Processor = $Processor.Name
        $computer.Ram = ([Math]::Round($ComputerSystem.TotalPhysicalMemory/1GB,0))
        $computer.Age = ([Math]::Round(((New-TimeSpan -Start ($OS.ConvertToDateTime($OS.InstallDate).ToShortDateString()) -End $(Get-Date)).Days / 365), 2))
        $computer.HardDriveSize = [Math]::Round($HardDrive.Size / 1GB, 0)
        $computer.HardDriveModel = $HardDrive.Model

        # Begin Evaluation

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

        if ($computer.Age -lt 5.0) {
            $computer.AgeValid = $true
        }

        if ($computer.HardDriveSize -ge 100) {
            $computer.HardDriveValid = $true
        }

        if ($computer.HardDriveModel -like "*SSD*") {
            $computer.SSD = $true
        }

        # Begin Output

        if ($computer.ArchitectureValid -and $computer.ProcessorValid -and $computer.RamValid -and $computer.AgeValid) {
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

        if ($computer.AgeValid) {
            Write-Host "$($computer.Age) years old" -ForegroundColor Green
        } else {
            Write-Host "$($computer.Age) years old" -ForegroundColor Red
        }

        
    } catch [System.UnauthorizedAccessException] {
        Write-Host "Access Denied" -ForegroundColor Red
        $computer.Error = "Access Denied"
    } catch [System.Runtime.InteropServices.COMException] {
        if (Test-Connection $RemoteComputerName -Count 1 -Quiet) {
            Write-Host "RPC server unavailable" -ForegroundColor Red
            $computer.Error = "RPC server unavailable"
        } else {
            Write-Host "Offline" -ForegroundColor Red
            $computer.Error = "Offline"
        }
    } catch [System.Management.ManagementException] {
        Write-Host "User credentials cannot be used for local connections" -ForegroundColor Red
        $computer.Error = "User credentials cannot be used for local connections"
    } catch {
        Write-Host "Unknown Error - $($_.Exception.Message) | $($_.Exception.GetType())" -ForegroundColor Red
        $computer.Error = "Unknown Error - $($_.Exception.Message) | $($_.Exception.GetType())"
    }

    $computers += $computer

}

Write-Host "Scan complete... Please select where to save the file..."

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "CSV (*.csv)|*.csv"
    FileName = "OSUpgradeEvaluation"
}
[void]$FileBrowser.ShowDialog()

$Path = $FileBrowser.FileName
$computers | Export-Csv -Path $Path

Write-Host "`n`nFile Saved... $($Path)" -ForegroundColor Green
Write-Host "Done!" -ForegroundColor Green