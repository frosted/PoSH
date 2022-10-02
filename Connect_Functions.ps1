
Function Connect-MECM
{    
    [CmdletBinding()]
    Param 
    (
        [Parameter(Mandatory=$false)][string]$ProviderMachineName=$Script:SMSProvider
    )
    Begin
    {
        Function Get-MECMSiteCode
        {

            Param 
            (
                [Parameter(Mandatory=$true)][string]$SMSProvider
            )

            $wqlQuery = “SELECT * FROM SMS_ProviderLocation”
            $a = Get-WmiObject -Query $wqlQuery -Namespace “root\sms” -ComputerName $SMSProvider
            $a | ForEach-Object {
                if($_.ProviderForLocalSite)
                    {
                        $script:SiteCode = $_.SiteCode
                    }
            }
            return $SiteCode
        }
    }
    Process
    {
        Try
        {
            # Site configuration
     
            $SiteCode = Get-MECMSiteCode -SMSProvider $ProviderMachineName

            # Customizations
            $initParams = @{}
            #$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
            #$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

            # Do not change anything below this line

            # Import the ConfigurationManager.psd1 module 
            if((Get-Module ConfigurationManager) -eq $null) {
                Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }

            # Connect to the site's drive if it is not already present
            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
                New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }

            # Set the current location to be the site code.
            Set-Location "$($SiteCode):\" @initParams
        }
        Catch
        {
            Write-Error "$($_.Exception.Message) - Line Number: $($_.InvocationInfo.ScriptLineNumber)"
        }
    }
    End
    {
        If ( (Get-Location | Select-Object -ExpandProperty Provider) -like '*CMSite*')
        {
            Write-Output 'Connection Successful'
        }
    }
}


