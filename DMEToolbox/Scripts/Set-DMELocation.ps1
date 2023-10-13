Function Set-DMELocation
{
    [CmdletBinding()]
    Param
    (
        [Parameter()]
        [ValidateSet('CMSite', 'FileSystem')]
        $Provider
    )

    Switch ($Provider)
    {
        'CMSite'
        {
            Try
            {
                If ((Get-Location).Provider.Name -ne 'CMSite')
                {
                    If (-not(Get-Module -Name ConfigurationManager))
                    {
                        Import-Module ConfigurationManager
                    }
                    else
                    {
                        Import-Module ConfigurationManager -Force
                    }
                    
                    Set-Location -Path "$((Get-PSDrive -PSProvider CMSite).Name):"
                }
            }
            Catch
            {
                Write-Error -Message "Unable to connect to CMSite.  Please be sure you're running this from a system with the CM Admin Console installed."
            }
        }

        'FileSystem'
        {
            # Check for log folder 
            If ((Get-Location).Provider.Name -ne 'FileSystem')
            {
                Set-Location -Path $env:SystemDrive
            }
        }
    }
}