## IN PROGRESS
Function Set-DMECollectionRefreshSettings
{
   [CmdletBinding(DefaultParameterSetName = 'Parameter Set 1', 
      SupportsShouldProcess = $true, 
      PositionalBinding = $false,
      HelpUri = 'http://www.microsoft.com/',
      ConfirmImpact = 'Medium')]
   Param
   (
      [Parameter(Mandatory = $True,
         ValueFromPipelineByPropertyName = $true,
         ParameterSetName = 'Parameter Set 1')]
      [PSObject]
      $RollbackObject,

      [Parameter(Mandatory = $True,
         ValueFromPipelineByPropertyName = $true,
         ParameterSetName = 'Parameter Set 1')]
      [string]
      $CollectionID,
      [Parameter(Mandatory = $False,
         ValueFromPipelineByPropertyName = $true,
         ParameterSetName = 'Parameter Set 1')]
      [string]
      $LimitingCollectionName,

      [Parameter(Mandatory = $False,
         ValueFromPipelineByPropertyName = $true,
         ParameterSetName = 'Parameter Set 1')]
      [string]
      $LimitingCollectionID,

      [Parameter(Mandatory = $False,
         ValueFromPipelineByPropertyName = $true,
         ParameterSetName = 'Parameter Set 1')]
      [string]
      $NewName,

      [Parameter(Mandatory = $False,
         ValueFromPipelineByPropertyName = $true,
         ParameterSetName = 'Parameter Set 1')]
      [ValidateSet('None', 'Manual', 'Periodic', 'Continuous', 'Both')]
      [string]
      $RefreshType,

      [Parameter(Mandatory = $False,
         ValueFromPipelineByPropertyName = $true,
         ParameterSetName = 'Parameter Set 1')]
      [string]
      $Comment,

      [string]
      $LogFolder
   )

   Begin
   {
      Set-DMELocation -Provider CMSite

      $Params = @{}
        
      If ($CollectionID) { $Params.Add('CollectionID', $CollectionID) }           
      If ($LimitingCollectionName) { $Params.Add('LimitingCollectionName', $LimitingCollectionName) }
      If ($LimitingCollectionID) { $Params.Add('LimitingCollectionID', $LimitingCollectionID) }
      If ($NewName) { $Params.Add('NewName', $NewName) }
      If ($RefreshType) { $Params.Add('RefreshType', $RefreshType) }
      If ($Comment) { $Params.Add('Comment', $Comment) }
      Write-Output $Params
   }
   Process
   {
      If ($pscmdlet.ShouldProcess("Target", "Operation"))
      {
         Set-CMCollection @Params 
         $RollbackObject | Export-Csv -Path "$($Global:logFolder)\Set-DMECollectionsByRefresh_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation -Append
      }
   }
   End
   {
   }
}




