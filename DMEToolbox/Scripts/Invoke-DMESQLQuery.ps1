<#
.Synopsis
    Run a SQL query using PowerShell
.DESCRIPTION
    Will open a connection to your SQL database, run a query and close the connection.  
.EXAMPLE
    Invoke-DMESQLQuery -SQLServer <servername> -SQLDatabase <databasename> -SQLQuery 'Select * From Table'
.Example
    Invoke-DMESQLQuery -SQLServer <servername> -SQLDatabase <databasename> -SQLQuery 'Select * From Table' -ExportToCSV 
.Example
    Invoke-DMESQLQuery -SQLServer <servername> -SQLDatabase <databasename> -SQLQuery 'Select * From Table' -ExportToCSV -ExportPath 'C:\Temp' -ExportName 'Invoke-SQLQueryResults.CSV'
.NOTES
    Author: Ed Frost
    Date:   2023.05.08
#>
Function Invoke-DMESQLQuery
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        $SQLServer,

        [string]
        $SQLDatabase,

        [string]
        $SQLUser,

        [securestring]
        $SQLPassword,

        [string]
        $SQLQuery,

        [switch]
        $ExportToCSV,

        [string]
        $ExportPath,

        [string]
        $ExportName,

        [string]
        [Validateset('AllTables', 'AllColumns', 'ReturnTable')]
        $StoredQuery,

        [string]
        $table
    )

    Begin
    {
        Function Connect-ToDB 
        {
            # define parameters
            param
            (
                [string]
                $ServerName,
                [string]
                $Database,
                [string]
                $SqlUser,
                [securestring]
                $SqlPassword
            )
            # create connection and save it as global variable
            $global:Connection = New-Object System.Data.SQLClient.SQLConnection

            # build param variable for splatting
            $Params = @{
                'trusted_connection'  = 'True'
                'integrated security' = 'False'
            }

            If ($servername)
            {
                $Params += @{
                    'server' = "'$servername'"
                }
            }

            If ($database)
            {
                $Params += @{
                    'database' = "'$database'"
                }
            }

            If ($sqluser)
            {
                $Params += @{
                    'user id' = "'$sqluser'"
                }
            }

            If ($sqlpassword)
            {
                $Params += @{
                    'Password' = "'$sqlpassword'"
                }
            }

            #$Connection.ConnectionString = $Params.Keys | ForEach-Object -Begin {'"'} -Process {"$_ = " + $Params.Item($_) + ";" } -End {'"'}


            $Connection.ConnectionString = "server='$servername';database='$database';trusted_connection=true"
            If ($sqluser -and $sqlpassword)
            {
                $Connection.ConnectionString += "; user id = '$sqluser'; Password = '$sqlpassword'; integrated security='False'"
            }
            $Connection.Open()
            Write-Verbose 'Connection established'
        }

        # function that executes sql commands against an existing Connection object; In pur case
        # the connection object is saved by the Connect-ToDB function as a global variable
        Function Invoke-SqlQuery 
        {
            # define parameters
            param
            (
                [string]
                $sqlquery
            )
            Begin 
            {
                If (!$Connection) 
                {
                    Throw "No connection to the database detected. Run command Connect-ToDB first."
                }
                elseif ($Connection.State -eq 'Closed') 
                {
                    Write-Verbose 'Connection to the database is closed. Re-opening connection...'
                    try 
                    {
                        # if connection was closed (by an error in the previous script) then try reopen it for this query
                        $Connection.Open()
                    }
                    catch 
                    {
                        Write-Verbose "Error re-opening connection. Removing connection variable."
                        Remove-Variable -Scope Global -Name Connection
                        throw "Unable to re-open connection to the database. Please reconnect using the Connect-ToDB commandlet. Error is $($_.exception)."
                    }
                }
            }
            Process 
            {
                #$Command = New-Object System.Data.SQLClient.SQLCommand
                $command = $Connection.CreateCommand()
                $command.CommandText = $sqlquery
                Write-Verbose "Running SQL query '$sqlquery'"
                try 
                {
                    $result = $command.ExecuteReader()
                }
                catch 
                {
                    $Connection.Close()
                }
                $Datatable = New-Object "System.Data.Datatable"
                $Datatable.Load($result)
                return $Datatable
            }
            End 
            {
                Write-Verbose "Finished running SQL query."
            }
        }

        Function Close-DBConnection
        {
            $Global:Connection.Close()
            Remove-Variable -Scope Global -Name Connection
        }

        Switch ($StoredQuery)
        {
            AllTables
            {
                ## Query: get all tables for CM_EDM database
                $SQLquery = "SELECT *
                FROM information_schema.tables
                WHERE table_catalog = '$($SQLDatabase)'"
            }
            AllColumns
            {
                ## Query: get column names for table
                $SQLquery = "exec sp_columns $($table)"   
            }
            ReturnTable
            {
                ## Query: get all data for table
                $SQLquery = "SELECT * FROM $($table)"   
            }
        }
    }
    Process
    {
        If (!$ExportPath) { $ExportPath = Split-Path -Path $MyInvocation.MyCommand.Definition }
        If ($ExportPath[-1] -eq '\') { $ExportPath = $ExportPath.TrimEnd('\') }
        If (!$ExportName) { $ExportName = $SQLDatabase + '_' + (Get-Date).ToString('MM-dd-yyy') + '.csv' }
        If ($SQLServer) { $params += @{ ServerName = $SQLServer } }
        If ($SQLDatabase) { $params += @{ Database = $SQLDatabase } }
        If ($SQLUser) { $params += @{ SQLUser = $SQLUser } }
        If ($SQLPassword) { $params += @{ SQLPassword = $SQLPassword } }
        
        Connect-ToDB @params
        $Result = Invoke-SqlQuery -sqlquery $SQLQuery    
    }
    End
    {
        Close-DBConnection

        If ($ExportToCSV)
        {
            $Result | Export-Csv -Path "$($ExportPath)\$($ExportName)" -Force
            Write-Host "Results exported to: $($ExportPath)\$($ExportName)"
        }
        Else
        {
            $Result
        }
    }
}