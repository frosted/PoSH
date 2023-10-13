<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER CollectionNames
Parameter description

.PARAMETER Type
Parameter description

.PARAMETER OutputFolder
Parameter description

.PARAMETER SelectCollections
Parameter description

.PARAMETER ReportOnly
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Start-DMEDeploymentsByCollection {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String[]]
        $CollectionNames,
        
        [Parameter()]
        [ValidateSet('Application', 'Update')]
        [String]
        $Type = 'Application',

        [Parameter()]
        [String]
        $OutputFolder = $env:TEMP,

        [Parameter()]
        [Switch]
        $SelectCollections,

        [Parameter()]
        [Switch]
        $ReportOnly
    )

    #$WhatIfPreference = $true
    If ($PSBoundParameters["SelectCollections"]) {
        $CollectionNames = Get-WmiObject -ComputerName $DMESiteServer -Class SMS_ApplicationAssignment -Namespace "root\sms\site_$($DMESiteCode)" -Filter "Enabled = 'false'" | Select-Object -ExpandProperty CollectionName -Unique | Out-GridView -Title "Select target collection(s)" -OutputMode Multiple
    }

    [System.Collections.Generic.List[psobject]]$Collections = @()
    $CollectionNames | ForEach-Object { $Collections += Get-WmiObject -ComputerName $DMESiteServer -Namespace "root\sms\site_$($DMESiteCode)" -Class SMS_Collection -Filter "Name = '$($_)'" }

    If ($CollectionNames.count -ne $Collections.count) {
        Foreach ($CollectionName in $Collections.Name) {
            Write-Host "'$($CollectionName)' " -NoNewline
            If ($CollectionName -in $Collections.Name) {
                Write-Host "Found" -ForegroundColor Green
            } else {
                Write-Host "Not Found" -ForegroundColor Red
            }
        }

        throw 'One or more collection(s) you entered could not be found. Please check your collection(s) and try again.'
        Exit
    }

    Switch ($Type) {
        'Application' {
            [System.Collections.Generic.List[psobject]]$ApplicationAssignments = @()
            foreach ($CollectionName in $Collections.Name) {
                $ApplicationAssignments += Get-WmiObject -ComputerName $DMESiteServer -Class SMS_ApplicationAssignment -Namespace "root\sms\site_$($DMESiteCode)" -Filter "CollectionName LIKE '%$($CollectionName)%' and Enabled = 'false'"
                $ApplicationAssignments | Export-Csv $OutputFolder\appAssignments_$CollectionName_$(Get-Date -Format 'yyyyMMddhhmm').csv
            }
            If (-Not($ReportOnly)) {
                foreach ($ApplicationAssignment in $ApplicationAssignments) {
                    $ApplicationAssignment.Enabled = $true
                    #$assignment.OverrideServiceWindows = $false
                    $ApplicationAssignment.Put()
                }
                foreach ($CollectionName in $Collections.Name) {
                    $ApplicationAssignments += Get-WmiObject -ComputerName $DMESiteServer -Class SMS_ApplicationAssignment -Namespace "root\sms\site_$($DMESiteCode)" -Filter "CollectionName like '%$($CollectionName)%'" -Verbose
                }
            } else {
                $ApplicationAssignments | Select-Object -Property ApplicationName, CollectionName, AssignmentName, OverrideServiceWindows, Enabled | Sort-Object -Property CollectionName, Enabled, ApplicationName
            }
        }
        'Update' {
            [System.Collections.Generic.List[psobject]]$UpdateDeployments = @()
            foreach ($CollectionID in $Collections.CollectionID) {
                $UpdateDeployments += Get-WmiObject -ComputerName $DMESiteServer -Class SMS_UpdatesAssignment -Namespace "root\sms\site_$($DMESiteCode)" -Filter "TargetCollectionID LIKE '%$($CollectionID)%' and Enabled = 'false'"
                $UpdateDeployments | Export-Csv $OutputFolder\updAssignments_$CollectionName_$(Get-Date -Format 'yyyyMMddhhmm').csv
            }
            If (-Not($ReportOnly)) {
                #$UpdateDeployments | ForEach-Object { Set-CMSoftwareUpdateDeployment -InputObject $_ -Enable $false -WhatIf:$WhatIfPreference }
                foreach ($UpdateDeployment in $UpdateDeployments) {
                    $UpdateDeployment.Enabled = $true
                    #$assignment.OverrideServiceWindows = $false
                    $UpdateDeployment.Put()
                }
            } else {
                $UpdateDeployments | Select-Object -Property AssignmentName, Enabled
            }
        }
    }
}