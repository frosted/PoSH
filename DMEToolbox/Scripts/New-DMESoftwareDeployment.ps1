# IN PROGRESS
Function New-DMESoftwareDeployment {
    [CmdLetBinding()]
    Param
    (
        [string]
        [Parameter(Mandatory = $true)]
        [ValidateSet("Application", "Update", "Baseline")]
        $Type,

        [string]
        [Parameter()]
        $LimitingCollection = 'Workstations | All',

        [string]
        [Parameter()]
        $DeploymentFolder, # = '.\DeviceCollection\CGI\#Deployments',
    
        [switch]
        [Parameter(Mandatory = $false)]
        $WhatIf
    )

    Set-Location $env:SystemDrive

    While (!($DeploymentTemplate)) {
        Write-Verbose -Message "Select file containing computers for direct membership"
        $DeploymentTemplate = Get-DMEFileOrFolder -Type File
    }

    # hoping to use this a list we can choose from where the new collections would be stored.  Took 29 minutes to complete in an environment with 167 containers and 4 sub levels
    #Measure-Command {
    #    $SubLevel = 0
    #    $CMCollectionPaths = [System.Collections.Generic.List[string]]::new()
    #    $CMCollectionPaths += Get-ChildItem -Path \DeviceCollection | Select-Object @{Name='Path';Expression={$_.PSChildName + '\' + $_.Name}},@{Name='SubLevel';Expression={$SubLevel}}
    #    Do {
    #        Write-Output "Sublevel $($SubLevel)"
    #        $CountBefore = $CMCollectionPaths.Count
    #        #$CMCollectionPaths | Where-Object SubLevel -eq $SubLevel | ForEach-Object { Get-ChildItem -Path $_.Path| Select-Object @{Name='Path';Expression={'DeviceCollection\' + $_.PSChildName + '\' + $_.Name}},@{Name='SubLevel';Expression={$SubLevel+1}} }
    #        $CMCollectionPaths | Where-Object SubLevel -eq $SubLevel | ForEach-Object { $CMCollectionPaths += Get-ChildItem -Path $_.Path| Select-Object @{Name='Path';Expression={$_.PSPath.Split('\',3)[-1] + '\' + $_.Name}},@{Name='SubLevel';Expression={$SubLevel+1}} }
    #        $CountAfter = $CMCollectionPaths.Count
    #        Write-Output "Added $($Countafter-$CountBefore) objects"
    #        $SubLevel = $SubLevel+1
    #    } While ($CountAfter -gt $CountBefore)
    #}
        
    Write-Verbose -Message "Using template located: $($DeploymentTemplate.FullName.ToString())"

    $LogDirectory = $DeploymentTemplate.Directory.ToString() + '\Logs'
    
    If (!(Test-Path $LogDirectory)) { New-Item -Path $LogDirectory -ItemType Directory }
    
    $LogFullPath = $LogDirectory + '\' + $DeploymentTemplate.BaseName.ToString() + $(If ($WhatIf) { "-WhatIf" }) + ".LOG"

    Write-Verbose "Logging to: $LogFullPath"
    
    Start-Transcript -Path $LogFullPath

    $Contents = Import-CSV $DeploymentTemplate.FullName | Where-Object CollName -ne $null

    $DateTimeFormat = "MM/dd/yyyy h:mm tt"

    Write-Verbose -Message "Please ensure the DateTime strings are in the format: $($DateTimeFormat)"

    Set-DMELocation -Provider CMSite
    
    If ( (Get-Location | Select-Object -ExpandProperty Provider) -notlike '*CMSite*' ) { Break }
    #$CMNameSpace = "root\sms\site_$DMESiteCode"
    #$CMConnection = ([wmiclass]("\\$DMESiteServer\ROOT\sms\site_" + $DMESiteCode + ":SMS_UpdatesAssignment")).CreateInstance()

    If ($Type -eq 'Application') {
        Write-Verbose -Message "Retrieving collections..."
        $Collections_All = Get-WmiObject -ComputerName $DMESiteServer -Namespace root\sms\site_$DMESiteCode -class sms_collection
    }

    foreach ($Content in $Contents) {
        If ($Type -like "Update") { 
            [boolean]$RestartServer = [System.Convert]::ToBoolean($Content.RestartSer)
            [boolean]$RestartWorkstation = [System.Convert]::ToBoolean($Content.Restartwrk)
        } Else { 
            $ApplicationName = $Content.Program 
            $DeployAction = $Content.DeployAction
        }
        $GroupName = $Content.SUName
        $CollName = $Content.CollName
        $DeploymentName = "$($GroupName)-$($Collname)"
        $DeployType = $Content.DeployType
        $AvailableDate = $Content.AvailableDate
        $AvailableTime = $Content.AvailableTime
        $DeadlineDate = $Content.DeadlineDate
        $DeadlineTime = $Content.DeadlineTime
        [boolean]$IgnoreMWforInstall = [System.Convert]::ToBoolean($Content.MWInstall)
        [boolean]$IgnoreMWforRestart = [System.Convert]::ToBoolean($Content.MWRestart)
        $AD = [DateTime]"$($AvailableDate) $($AvailableTime)"
        $DD = [DateTime]"$($DeadlineDate) $($DeadlineTime)"
        #$AD=[DateTime]::ParseExact("$($AvailableDate) $($AvailableTime)",$DateTimeFormat,[Globalization.CultureInfo]::CreateSpecificCulture('en-CA'))
        #$DD=[DateTime]::ParseExact("$($DeadlineDate) $($DeadlineTime)",$DateTimeFormat,[Globalization.CultureInfo]::CreateSpecificCulture('en-CA'))
        $UserNotification = $Content.UserNotification
        
        Switch ($Type) {
            'Update' {
                #Software Update Deployment 
                $DeploymentExist = $False

                #Get the deployments
                $Deployments = Get-WmiObject -ComputerName $DMESiteServer -Namespace root\sms\site_$($DMESiteCode) -Class SMS_UpdatesAssignment -Filter "AssignmentName=""$DeploymentName"""

                #Check if the deployment name already exist or not
                if ($Deployments.AssignmentName) {
                    Write-Warning -Message "[INFO]`t Deployment $DeploymentName already exist, modifying deployment" 
                    Write-Verbose "Deployment $DeploymentName already exist, modifying deployment" 
                    $DeploymentExist = $True
                }

                Write-Verbose "-SoftwareUpdateGroupName $GroupName -CollectionName $CollName -DeploymentName $DeploymentName"
                Write-Verbose "-DeploymentType $DeployType -SendWakeUpPacket $False -VerbosityLevel AllMessages"
                Write-Verbose "-TimeBasedOn LocalTime -DeploymentAvailableDay $AvailableDate -DeploymentAvailableTime $AvailableTime -EnforcementDeadlineDay $DeadlineDate"
                Write-Verbose "-EnforcementDeadline $DeadlineTime -UserNotification DisplayAll -SoftwareInstallation $IgnoreMWforInstall -AllowRestart $IgnoreMWforRestart -RestartServer $RestartServer"
                Write-Verbose "-RestartWorkstation $RestartWorkstation -PersistOnWriteFilterDevice $False -GenerateSuccessAlert $False "
                Write-Verbose "-DisableOperationsManagerAlert $True -GenerateOperationsManagerAlert $False -ProtectedType NoInstall -UnprotectedType NoInstall"
                Write-Verbose "-UseBranchCache $False -DownloadFromMicrosoftUpdate $False -AllowUseMeteredNetwork $True"

                If (Get-CMSoftwareUpdateGroup -Name $GroupName) {
                    If ($WhatIf) {   
                        If (!$DeploymentExist) { 
                            New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $GroupName -CollectionName $CollName `
                                -DeploymentName $DeploymentName -DeploymentType $DeployType -SendWakeUpPacket $False -VerbosityLevel AllMessages `
                                -TimeBasedOn LocalTime -AvailableDateTime $AD  -DeadlineDateTime $DD `
                                -UserNotification $UserNotification -SoftwareInstallation $IgnoreMWforInstall -AllowRestart $IgnoreMWforRestart `
                                -RestartServer $RestartServer -RestartWorkstation $RestartWorkstation -PersistOnWriteFilterDevice $False -GenerateSuccessAlert $false `
                                -DisableOperationsManagerAlert $false -GenerateOperationsManagerAlert $false -ProtectedType RemoteDistributionPoint -UseBranchCache $false `
                                -DownloadFromMicrosoftUpdate $false -UseMeteredNetwork $false -RequirePostRebootFullScan $True -WhatIf
                        } Else {
                            Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $GroupName -CollectionName $CollName `
                                -DeploymentName $DeploymentName -DeploymentType $DeployType -SendWakeUpPacket $False -VerbosityLevel AllMessages `
                                -TimeBasedOn LocalTime -AvailableDateTime $AD -DeploymentExpireDateTime $DD `
                                -UserNotification $UserNotification -SoftwareInstallation $IgnoreMWforInstall -AllowRestart $IgnoreMWforRestart `
                                -RestartServer $RestartServer -RestartWorkstation $RestartWorkstation -PersistOnWriteFilterDevice $False -GenerateSuccessAlert $false `
                                -DisableOperationsManagerAlert $false -GenerateOperationsManagerAlert $false -ProtectedType RemoteDistributionPoint -UseBranchCache $false `
                                -DownloadFromMicrosoftUpdate $false -AllowUseMeteredNetwork $false -RequirePostRebootFullScan $True -WhatIf
                        }
                    } Else {
                        If (!$DeploymentExist) {
                            New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $GroupName -CollectionName $CollName `
                                -DeploymentName $DeploymentName -DeploymentType $DeployType -SendWakeUpPacket $False -VerbosityLevel AllMessages `
                                -TimeBasedOn LocalTime -AvailableDateTime $AD  -DeadlineDateTime $DD `
                                -UserNotification $UserNotification -SoftwareInstallation $IgnoreMWforInstall -AllowRestart $IgnoreMWforRestart `
                                -RestartServer $RestartServer -RestartWorkstation $RestartWorkstation -PersistOnWriteFilterDevice $False -GenerateSuccessAlert $false `
                                -DisableOperationsManagerAlert $false -GenerateOperationsManagerAlert $false -ProtectedType RemoteDistributionPoint -UseBranchCache $false `
                                -DownloadFromMicrosoftUpdate $false -UseMeteredNetwork $false -RequirePostRebootFullScan $True
                        } Else {
                            Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $GroupName -CollectionName $CollName `
                                -DeploymentName $DeploymentName -DeploymentType $DeployType -SendWakeUpPacket $False -VerbosityLevel AllMessages `
                                -TimeBasedOn LocalTime -AvailableDateTime $AD  -DeploymentExpireDateTime $DD `
                                -UserNotification $UserNotification -SoftwareInstallation $IgnoreMWforInstall -AllowRestart $IgnoreMWforRestart `
                                -RestartServer $RestartServer -RestartWorkstation $RestartWorkstation -PersistOnWriteFilterDevice $False -GenerateSuccessAlert $false `
                                -DisableOperationsManagerAlert $false -GenerateOperationsManagerAlert $false -ProtectedType RemoteDistributionPoint -UseBranchCache $false `
                                -DownloadFromMicrosoftUpdate $false -AllowUseMeteredNetwork $false -RequirePostRebootFullScan $True 
                        }
                    }
                    if ($error) {
                        Write-Output "[INFO]`t Could not create SU Deployment $($DeploymentName) ,Please check further"
                        Write-Verbose "$date [INFO]`t Could not SU Deployment $($DeploymentName) ,Please check further: $error"
                        $error.Clear()
                    } else {                    
                        Write-Output "[INFO]`t Created SU Deployment : $($DeploymentName)"
                        Write-Verbose "$date [INFO]`t Created SU Deployment : $($DeploymentName)"
                    }
                }
            }

            'Application' {
                #Application Deployment
                Write-Verbose "$($Content.ApplicationID) $($ApplicationName) $($CollName) $($AD)"

                If ($WhatIf) {
                    If (!( $Collections_All | Where-Object Name -eq $CollName)) {
                        Write-Verbose "$($CollName) not found.  Creating Collection."
                        New-CMCollection -CollectionType Device -LimitingCollectionName $LimitingCollection -Name $CollName -WhatIf
                    }
                    New-CMApplicationDeployment -CollectionName $CollName -Name $ApplicationName -TimeBaseOn LocalTime -AvailableDateTime $AD -DeadlineDateTime $DD -DeployPurpose $DeployType -DeployAction $DeployAction -OverrideServiceWindow $IgnoreMWforInstall -RebootOutsideServiceWindow $IgnoreMWforRestart -UserNotification $UserNotification -WhatIf
                } Else {
                    If (!( $Collections_All | Where-Object Name -eq $CollName)) {
                        Write-Verbose "$($CollName) not found.  Creating Collection."
                        New-CMCollection -CollectionType Device -LimitingCollectionName $LimitingCollection -Name $CollName 
                        Start-Sleep -Seconds 5
                    }

                    If (Get-CMApplicationDeployment -CollectionName $CollName -Name $ApplicationName) {
                        Set-CMApplicationDeployment -CollectionName $CollName -ApplicationName $ApplicationName -AvailableDateTime $AD -DeadlineDateTime $DD -OverrideServiceWindow $IgnoreMWforInstall -RebootOutsideServiceWindow $IgnoreMWforRestart -UserNotification $UserNotification
                    } Else {
                        New-CMApplicationDeployment -CollectionName $CollName -Name $ApplicationName -TimeBaseOn LocalTime -AvailableDateTime $AD -DeadlineDateTime $DD -DeployPurpose $DeployType -DeployAction $DeployAction -OverrideServiceWindow $IgnoreMWforInstall -RebootOutsideServiceWindow $IgnoreMWforRestart -UserNotification $UserNotification
                    }
                }

                If ($error) {
                    Write-Verbose "[INFO]`t Could not create Deployment $($Content.Program) for $($CollName) ,Please check further"
                    Write-Verbose "Could not create Deployment  $($Content.Program) for $($CollName), Please check further: $($error)"
                    $error.Clear()
                } Else {                    
                    Write-Verbose "[INFO]`t Created Deployment $($Content.Program) for $($CollName)"
                    Write-Verbose "Created Deployment  $($Content.Program) for $($CollName)"
                }
                
                If ($null -ne $DeploymentFolder) {
                    Move-CMObject -FolderPath $DeploymentFolder -InputObject $(Get-CMCollection -Name $Content.CollName)
                }
            }

            'Baseline' {
                #Write-Host "Baseline"
                If ((Get-CMCollection -Name $CollName)) {
                    If ($WhatIf) {
                        $sched = New-CMSchedule -Start $dd -RecurInterval Days -RecurCount 1 
                        New-CMBaselineDeployment -CollectionName $CollName -Name $ApplicationName -Schedule $sched -WhatIf 
                        
                    } Else {
                        $sched = New-CMSchedule -Start $dd -RecurInterval Days -RecurCount 1 
                        New-CMBaselineDeployment -CollectionName $CollName -Name $ApplicationName -Schedule $sched 
                    }

                    If ($error) {
                        Write-Verbose "[INFO]`t Could not create Deployment $($Content.Program) for $($CollName) ,Please check further"
                        Write-Verbose "Could not create Deployment  $($Content.Program) for $($CollName), Please check further: $($error)"
                        $error.Clear()
                    } else {                    
                        Write-Verbose "[INFO]`t Created Deployment $($Content.Program) for $($CollName)"
                        Write-Verbose "Created Deployment  $($Content.Program) for $($CollName)"
                    }
                } Else {
                    Write-Verbose "Collection $($CollName) does not exist.  Please check collection name and try again."
                }
            }
        }  
    }	
}