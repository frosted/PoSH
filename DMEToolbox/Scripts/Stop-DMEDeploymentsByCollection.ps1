<#
.SYNOPSIS
Stop all deployments assigned to a collection

.DESCRIPTION
This function will allow you to specify or choose a collection, and stop all active deployments.  

.PARAMETER CollectionNames
Specify one or more collections

.PARAMETER Type
Specify deployment type, Application or Update.  Input for this parameter will be validated at runtime.

.PARAMETER OutputFolder
Defaulting to temp folder, you may specify where the CSV file will be stored.  This CSV file will capture the current state of the deployments for restore.

.PARAMETER ReportOnly
Similar to whatif, use this switch to only report on what deployments will be targeted.

.EXAMPLE
# Stop all update deployments for 'Executive Systems' and 'Retail Systems' collections
Stop-DMEDeploymentsByCollection -Type Updates -CollectionNames @('Executive Systems','Retail Systems')

.EXAMPLE
# Report on all application deployments assigned to the 'All Workstations' collection, but do not stop them yet.
Stop-DMEDeploymentsByCollection -Type Application -CollectionNames @('All Workstations') -ReportOnly

.NOTES
General notes
#>
Function Stop-DMEDeploymentsByCollection {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String[]]
        $CollectionNames = @('deploy 12 - executive ous', 'deploy 11 - support exec ous'),
        
        [Parameter()]
        [ValidateSet('Application', 'Update')]
        [String]
        $Type = 'Application',

        [Parameter()]
        [Switch]
        $SelectCollections,

        [Parameter()]
        [String]
        $OutputFolder = $env:TEMP,

        [Parameter()]
        [Switch]
        $ReportOnly
    )

    #$WhatIfPreference = $true
    If ($PSBoundParameters["SelectCollections"]) {
        $CollectionNames = Get-WmiObject -ComputerName $DMESiteServer -Class SMS_ApplicationAssignment -Namespace "root\sms\site_$($DMESiteCode)" -Filter "Enabled = 'true'" | Select-Object -ExpandProperty CollectionName -Unique | Out-GridView -Title "Select target collection(s)" -OutputMode Multiple
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
                $ApplicationAssignments += Get-WmiObject -ComputerName $DMESiteServer -Class SMS_ApplicationAssignment -Namespace "root\sms\site_$($DMESiteCode)" -Filter "CollectionName LIKE '%$($CollectionName)%' and Enabled = 'true'"
                $ApplicationAssignments | Export-Csv $OutputFolder\appAssignments_$CollectionName_$(Get-Date -Format 'yyyyMMddhhmm').csv
            }
            If (-Not($ReportOnly)) {
                foreach ($ApplicationAssignment in $ApplicationAssignments) {
                    $ApplicationAssignment.Enabled = $false
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
                $UpdateDeployments += Get-WmiObject -ComputerName $DMESiteServer -Class SMS_UpdatesAssignment -Namespace "root\sms\site_$($DMESiteCode)" -Filter "TargetCollectionID LIKE '%$($CollectionID)%' and Enabled = 'true'"
                $UpdateDeployments | Export-Csv $OutputFolder\updAssignments_$CollectionName_$(Get-Date -Format 'yyyyMMddhhmm').csv
            }
            If (-Not($ReportOnly)) {
                #$UpdateDeployments | ForEach-Object { Set-CMSoftwareUpdateDeployment -InputObject $_ -Enable $false -WhatIf:$WhatIfPreference }
                foreach ($UpdateDeployment in $UpdateDeployments) {
                    $UpdateDeployment.Enabled = $false
                    #$assignment.OverrideServiceWindows = $false
                    $UpdateDeployment.Put()
                }
            } else {
                $UpdateDeployments | Select-Object -Property AssignmentName, Enabled
            }
        }
    }
}