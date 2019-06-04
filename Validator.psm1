function Validate-Computer {
    param (
        $computer
    )
    
    $computer | Add-Member -MemberType NoteProperty -Name OSCurrent -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name ArchitectureValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name ProcessorValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name RamValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name AgeValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name HardDriveValid -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name SSD -Value $false
    $computer | Add-Member -MemberType NoteProperty -Name SuggestedAction -Value 0
    
    $computer.OSCurrent = $computer.OS -like "10.*"
    $computer.ArchitectureValid = $computer.Architecture -like "*64-bit*"
    $computer.RamValid = ([int]$computer.Ram) -ge 8
    $computer.AgeValid = ([int]$computer.Age) -le 5.0
    $computer.HardDriveValid = ([int]$computer.HardDriveSize) -ge 100
    $computer.SSD = ($computer.HardDriveModel -like "*SSD*" -or $computer.HardDriveModel -like "*Virtual*")
    
    if ($computer.Processor -like "*Core*") {
        if ($computer.Processor.Substring(18,1) -like "i") {
            if ($computer.Processor.Substring(19,1) -eq 5) {
                if ($computer.Processor.Substring(21,4) -ge 3000) {
                    $computer.ProcessorValid = $true
                }
            } elseif ($computer.Processor.Substring(19,1) -gt 5) {
                $computer.ProcessorValid = $true
            }
        }
    } elseif ($computer.Processor -like "*Xeon*") {
        $computer.ProcessorValid = $true
    }

    $partsNeeded = !($computer.RamValid -and 
        $computer.HardDriveValid -and 
        $computer.SSD)
    if ($computer.OSCurrent) {
        if ($partsNeeded) {
            # Replace Parts
            $computer.SuggestedAction = 2
        } else {
            # OK
            $computer.SuggestedAction = 0
        }
    } else {
        if ($partsNeeded) {
            # Upgrade OS & Replace Parts
            $computer.SuggestedAction = 3
        } else {
            # Upgrade OS
            $computer.SuggestedAction = 1
        }
    } 
    
    if (!($computer.ProcessorValid -and $computer.ArchitectureValid -and $computer.AgeValid)) {
        # Replace Computer
        $computer.SuggestedAction = -1
    }

    return $computer
}

Export-ModuleMember -Function Validate-Computer