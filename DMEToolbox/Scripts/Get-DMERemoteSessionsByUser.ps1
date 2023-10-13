Function Get-DMERemoteSessionsByUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String[]]
        $Hostnames,
        [Parameter()]
        [String[]]
        $Usernames = @($env:USERNAME),
        [Parameter()]
        [Switch]
        $Kill
    )

    [system.Collections.Generic.List[string]]$results = @()
    foreach ($Hostname in $Hostnames) {
        foreach ($Username in $Usernames) {
            $queryResult = query user $Username /server:$Hostname 2>&1 

            If ($null -eq $queryResult) {
                $queryResult = $Hostname + ": Falied to connect to $Host"
            } Else {
                $queryResult = $Hostname + ': ' + $queryResult
                Write-Progress -Activity "Checking for remote RD session for $Username on $Hostname" -Status $queryResult.ToString()  -PercentComplete (($Hostnames.IndexOf($Hostname)) / ($Hostnames.count) * 100) 
            }
            $results.Add($queryResult)
        }
    }

    If ($Kill) {
        Write-Warning 'The option to kill your remote sessions is coming soon.'
    }

    Return $results
}