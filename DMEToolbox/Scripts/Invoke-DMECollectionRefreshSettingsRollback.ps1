# IN PROGRESS
Function Invoke-DMECollectionRefreshSettingsRollback
{
    [CmdletBinding(DefaultParameterSetName = 'Parameter Set 1', 
        SupportsShouldProcess = $true, 
        PositionalBinding = $false,
        HelpUri = 'http://www.microsoft.com/',
        ConfirmImpact = 'Medium')]
    Param
    (
        #[Parameter(Mandatory = $True,
        #    ValueFromPipelineByPropertyName = $true,
        #    ParameterSetName = 'Parameter Set 1')]
        #[string]
        #$ModuleDirectory
    )

    Begin
    {
        $rollbackFile = Get-DMEFile -InitialDirectory $InitialDirectory
        
        $rollbackTargets = Import-Csv -Path $rollbackFile

        Write-Verbose "WhatIf Preference: $($WhatIfPreference)"

        Set-DMELocation -Provider CMSite
    }
    Process
    {
        # Check rollback settings
        
        
        # Remove configurationmanager module dependency
        $rollbackTargets | ForEach-Object { Write-Verbose "$_.CollectionID"; Set-CMCollection -CollectionId $_.CollectionID -RefreshType $_.RefreshType -LimitingCollectionId $_.LimitToCollectionID -WhatIf:$WhatIfPreference }
    }
    End
    {
        Write-Verbose -Message "Rollback complete."
    }
}