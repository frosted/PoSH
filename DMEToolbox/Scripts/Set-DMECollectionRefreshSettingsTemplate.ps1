# IN PROGRESS
Function Set-DMECollectionRefreshSettingsTemplate
{
    [CmdletBinding(DefaultParameterSetName = 'Parameter Set 1', 
        SupportsShouldProcess = $true, 
        PositionalBinding = $false,
        HelpUri = 'http://www.microsoft.com/',
        ConfirmImpact = 'Medium')]
    Param
    (
        [Parameter(Mandatory = $False,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'Parameter Set 1')]
        [string]
        $InitialDirectory = $env:USERPROFILE
    )

    Begin
    {
        $TemplateFile = Get-DMEFileOrFolder -InitialDirectory $InitialDirectory -Type File
        
        $CollectionTargets = Import-Csv -Path $TemplateFile.FullName | Where-Object { $_.Action -NotIn @('None', 'Review') }

        Write-Verbose "WhatIf Preference: $($WhatIfPreference)"

        Set-DMELocation -Provider CMSite

        Function Convert-CollectionSchedule
        {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory)]
                [String]
                $ScheduleString
            )
            Switch ($ScheduleString)
            {
                'Every 1 days'
                {
                    $Schedule = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Days -RecurCount 1
                }
                'Every 7 days'
                {
                    $Schedule = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Days -RecurCount 7
                }
                'Every 1 hour'
                {
                    $Schedule = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Hours -RecurCount 1
                }
                'Every 4 hours'
                {
                    $Schedule = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Hours -RecurCount 4
                }
            }
            Return $Schedule
        }
    }
    Process
    {
        # Remove configurationmanager module dependency
        $CollectionTargets | ForEach-Object { Write-Verbose "$($_.CollectionID)"; Set-CMCollection -CollectionId $_.CollectionID -RefreshSchedule $(Convert-CollectionSchedule -ScheduleString $_.'Update Schedule') -RefreshType $_.RefreshType -LimitingCollectionId $_.LimitToCollectionID -WhatIf:$WhatIfPreference }
    }
    End
    {
        Write-Verbose -Message "Refresh settings deployment completed on $($CollectionTargets.count) collections."
    }
}