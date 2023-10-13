<#
.SYNOPSIS
Creates global variables to be referenced within this module.

.DESCRIPTION
This function references the values defined in a config file (DMEEnvironment.conf, by default), and uses the 
name-value pairs to create global variables - referenced throughout this module

.PARAMETER Config
Full path to the config file containing the name-value pairs to be imported

.EXAMPLE
Initialize-DMEEnvironment -Config c:\DMEEnvironment.conf

.NOTES
Author: Ed Frost
Date:   2023.08.23
#>
Function Initialize-DMEEnvironment {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $False, HelpMessage = "Specify the full path to the config file.")]
        [ValidateScript({ If ($_) { Test-Path -Path $_ -ErrorAction SilentlyContinue } })]
        [string]$Config,
        [Parameter(Mandatory = $False)]
        [string]$SiteServer,
        [Parameter(Mandatory = $False)]
        [string]$SiteCode,
        [Parameter(Mandatory = $False)]
        [string]$SiteDBServer,
        [Parameter(Mandatory = $False)]
        [string]$SiteDB,
        [Parameter(Mandatory = $False)]
        [string]$Proxy,
        [Parameter(Mandatory = $False)]
        [string]$ModulePath,
        [Parameter(Mandatory = $False)]
        [string]$LogPath,
        [Parameter(Mandatory = $False)]
        [switch]$SkipVerification
    )

    Begin {
        If ($Config) {
            Set-Variable -Name DMEConfigPath $Config -Scope global
            
            foreach ($Line in $(Get-Content $Config)) {
                #Write-Verbose -Message $Line
                Set-Variable -Name $Line.Split('=')[0] -Value ($Line.Split('=', 2)[1]).Trim() -Scope global
            }
        }

        If ($SiteServer  ) { $Global:DMESiteServer = $SiteServer }  
        If ($SiteCode    ) { $Global:DMESiteCode = $SiteCode }
        If ($SiteDBServer) { $Global:DMESiteDBServer = $SiteDBServer }
        If ($SiteDB      ) { $Global:DMESiteDB = $SiteDB }
        If ($Proxy       ) { $Global:DMEProxy = $Proxy }
        If ($ModulePath  ) { $Global:DMEModulePath = $ModulePath }

        Write-verbose $("Global Variables" + [char]10 +
            "Site Server..$Global:DMESiteServer" + [char]10 +
            "Site Code....$Global:DMESiteCode" + [char]10 +   
            "DB Server....$Global:DMESiteDBServer" + [char]10 +
            "DB...........$Global:DMESiteDB" + [char]10 +
            "Proxy........$Global:DMEProxy" + [char]10 +      
            "Module Path..$Global:DMEModulePath")
    }

    Process {
        If ($PSBoundParameters.ContainsKey('SkipVerification')) {
            Write-Verbose 'Skipping verification phase.'
        } else {
            if ($null -ne $SiteServer) {
                Write-Verbose -Message "Site Server: $($SiteServer)" 
                try {
                    Write-Verbose -Message "Testing connection with $($SiteServer)"
                    Test-NetConnection -ComputerName $SiteServer -InformationLevel Quiet -ErrorAction Stop -WarningAction Stop | Out-Null
                } catch [System.UnauthorizedAccessException] {
                    Write-Warning -Message "Access denied" ; break
                } catch [System.Exception] {
                    Write-Warning -Message "Unable to connect to Site Server $(SiteServer).  Please check your values in $($Config)."; break
                }
            }

            if ($null -ne $DMESiteCode) {
                Write-Verbose -Message "Site Code: $($DMESiteCode)"
            } else {
                # Determine SiteCode from WMI
                try {
                    Write-Verbose -Message "Determining Site Code for Site server: '$($DMESiteServer)'"
                    $SiteCodeObjects = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $DMESiteServer -ErrorAction Stop
                    foreach ($SiteCodeObject in $SiteCodeObjects) {
                        if ($SiteCodeObject.ProviderForLocalSite -eq $true) {
                            $DMESiteCode = $SiteCodeObject.SiteCode
                            Write-Verbose -Message "Site Code: $($DMESiteCode)"
                        }
                    }
                } catch [System.UnauthorizedAccessException] {
                    Write-Warning -Message "Access denied" ; break
                } catch [System.Exception] {
                    Write-Warning -Message "Unable to determine Site Code.  Please check your values in $($Config)."; break
                }
            }
        }
        
    }

    End {
        Write-Verbose -Message "Initialization complete."
    }
}