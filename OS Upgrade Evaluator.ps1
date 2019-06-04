# Init Functions

$ScanComputer = {
    param (
        $RemoteComputerName,
        $credentials,
        $UseCurrentCredentials
    )

    $localComputerName = (Get-WmiObject Win32_ComputerSystem).Name

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
    $computer | Add-Member -MemberType NoteProperty -Name Error -Value $null
    
    try {

        if ($RemoteComputerName -eq $localComputerName) {
            $ComputerSystem = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
            $OS = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
            $Processor = Get-WmiObject Win32_processor -ErrorAction Stop
            $HardDrive = (Get-WmiObject Win32_DiskDrive -ErrorAction Stop) | Where-Object DeviceID -Like "*PhysicalDrive0*"
        } elseif (!$UseCurrentCredentials) {
            $ComputerSystem = Get-WmiObject Win32_ComputerSystem -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop
            $OS = Get-WmiObject Win32_OperatingSystem -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop
            $Processor = Get-WmiObject Win32_processor -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop
            $HardDrive = (Get-WmiObject Win32_DiskDrive -Impersonation 3 -ComputerName $RemoteComputerName -Credential $credentials -ErrorAction Stop) | Where-Object DeviceID -Like "*PhysicalDrive0*"
        } else {
            $ComputerSystem = Get-WmiObject Win32_ComputerSystem -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop
            $OS = Get-WmiObject Win32_OperatingSystem -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop
            $Processor = Get-WmiObject Win32_processor -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop
            $HardDrive = (Get-WmiObject Win32_DiskDrive -Impersonation 3 -ComputerName $RemoteComputerName -ErrorAction Stop) | Where-Object DeviceID -Like "*PhysicalDrive0*"
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

        Write-Host "[Done]" -ForegroundColor Green
        
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

    return $computer
    
}

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

foreach ($RemoteComputer in $networkComputers) {
    $RemoteComputerName = $RemoteComputer.Path.Split("/")[3]
    Start-Job -ScriptBlock $ScanComputer -Name $RemoteComputerName -ArgumentList $RemoteComputerName, $credentials, $UseCurrentCredentials
}

$jobs = get-job
do {
    Clear-Host
    $jobsRunning = 0
    $jobsFailed = 0
    $jobsCompleted = 0
    foreach ($job in $jobs) {
        if ($job.state -like "*Running*") {
            $jobsRunning++
            Write-Host "$($job.state) scan on $($job.name)" -ForegroundColor Yellow
        } elseif ($job.state -like "*Blocked*") {
            $jobsFailed++
            Write-Host "$($job.state) scan on $($job.name)" -ForegroundColor Red
        } else {
            $jobsCompleted++
            Write-Host "$($job.state) scan on $($job.name)" -ForegroundColor Green
        }
    }

    Write-Host "`n`n$($jobsRunning) scans running" -ForegroundColor Yellow
    Write-Host "$($jobsFailed) scans failed" -ForegroundColor Red
    Write-Host "$($jobsCompleted) scans completed" -ForegroundColor Green
    Start-Sleep 1
} while ($jobsRunning -gt 0)

foreach ($job in $jobs) {
    $computers += $job | Receive-Job
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

