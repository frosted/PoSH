<#
.SYNOPSIS
Provides deployment status for applications.  

.DESCRIPTION
Returns the deployment status for a specific application

.PARAMETER DBServer
Name or instance of CM DB Server

.PARAMETER DBName
Name of database

.PARAMETER ApplicationName
Name of application.

.EXAMPLE
$objDeployment = Get-DMEDeploymentReport -DBServer dbserver.local -DBName cm_lab -ApplicationName PBIDesktopSetup_x64

.NOTES
2023.08.04 - added to DMEToolbox

Feature request
-choose from available deployments
-filter to collection
#>
Function Get-DMEDeploymentReport {
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory = $False, HelpMessage = "Specify application display name.")]
        [string]$ApplicationName,
        [parameter(Mandatory = $False, HelpMessage = "Use this switch to bring up a list of applications to choose from")]
        [switch]$SelectApplication
    )

    Begin {
        # Select application
        if ($PSBoundParameters["SelectApplication"]) {
            try {
                $ApplicationName = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery `
                    "SELECT Manufacturer, DisplayName, NumberOfDeployments, DateCreated, CreatedBy 
                    FROM dbo.fn_ListApplicationCIs(1033) 
                    WHERE IsLatest = 'True' AND IsEnabled = 'True' 
                    AND IsExpired = 'False' 
                    AND NumberOfDeployments > 0
                    ORDER BY Manufacturer, DisplayName" | Out-GridView -Title 'Select Application Title' -OutputMode Single | Select-Object -ExpandProperty DisplayName 
            } catch [system.UnauthorizedAccessException] {
                Write-Warning -Message "Access Denied."
            } catch [system.Exception] {
                Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to $($DMESiteDBServer) or check your values in your config file."; break
            }
        } 

        $ApplicationNames = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery "SELECT DISTINCT DisplayName FROM dbo.fn_ListApplicationCIs(1033)"
        
        # Verify application name exist
        If (-Not($ApplicationNames | Where-Object DisplayName -eq $ApplicationName)) {
            Write-Warning -Message "$($ApplicationName) could not be found. Check input and try again."; break
        }

        $query = "SELECT distinct
            vrs.Name0 [Computer Name], vgos.Caption0 [OS],vrs.User_Name0 [User Name],
            IIf([EnforcementState]=1001,'Installation Success',
            IIf([EnforcementState]>=1000 And [EnforcementState]<2000 And [EnforcementState]<>1001,'Installation Success',
            IIf([EnforcementState]>=2000 And [EnforcementState]<3000,'In Progress', 
            IIf([EnforcementState]>=3000 And [EnforcementState]<4000,'Requirements Not Met', 
            IIf([EnforcementState]>=4000 And [EnforcementState]<5000,'Unknown', 
            IIf([EnforcementState]>=5000 And [EnforcementState]<6000,'Error','Unknown')))))) AS Status
            FROM dbo.v_R_System AS vrs
            INNER JOIN (dbo.vAppDeploymentResultsPerClient
            INNER JOIN v_CIAssignment
            ON dbo.vAppDeploymentResultsPerClient.AssignmentID = v_CIAssignment.AssignmentID)
            ON vrs.ResourceID = dbo.vAppDeploymentResultsPerClient.ResourceID
            INNER JOIN dbo.fn_ListApplicationCIs(1033) lac
            ON lac.ci_id=dbo.vAppDeploymentResultsPerClient.CI_ID
            INNER JOIN dbo.v_GS_WORKSTATION_STATUS AS vgws
            ON vgws.ResourceID=vrs.resourceid
            INNER JOIN v_FullCollectionMembership coll
            ON coll.ResourceID = vrs.ResourceID
            INNER JOIN dbo.v_GS_OPERATING_SYSTEM AS vgos
            ON vgos.ResourceID = vrs.ResourceID
            WHERE lac.DisplayName = '$($ApplicationName)'
            --and CollectionName = 'provide your collection name here'
            "
    }

    Process {
        try {
            $return = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $query  
        } catch [system.UnauthorizedAccessException] {
            Write-Warning -Message "Access Denied."
        } catch [system.Exception] {
            Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to $($DMESiteDBServer) or check your values in your config file."; break
        }
        
        $Deployment = [PSCustomObject]@{
            Application = $ApplicationName
            Success     = ($return | Where-Object Status -eq 'Installation Success' | Measure-Object).Count
            Installing  = ($return | Where-Object Status -eq 'In Progress' | Measure-Object).Count
            ReqNotMet   = ($return | Where-Object Status -eq 'Requirements Not Met' | Measure-Object).Count
            Error       = ($return | Where-Object Status -eq 'Error' | Measure-Object).Count
            Unknown     = ($return | Where-Object Status -eq 'Unknown' | Measure-Object).Count
            Targeted    = $return.count
            Compliance  = If ($return.count -gt 0) { ($return | Where-Object Status -eq 'Installation Success' | Measure-Object).Count / $return.count } Else { 0 }
        }
    }

    End {
        Return $Deployment
    }
}