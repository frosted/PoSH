


$DMESCOMServer = 'btwp001642.corp.ads'

$ModulePath = "\\$DMESCOMServer\d$\Program Files\Microsoft System Center 2016\Operations Manager\Powershell\"

Test-Path $ModulePath

Copy-Item -Path $ModulePath -Destination "$DMEModulePath\OperationsManager" -Container -Recurse

$SDKPath = "\\$DMESCOMServer\d$\Program Files\Microsoft System Center 2016\Operations Manager\Server\SDK Binaries\"

Test-Path $SDKPath

Get-ChildItem $SDKPath | ForEach-Object { Copy-Item -Path $_.FullName -Destination $Env:windir\assembly }

$NewPSPath = [System.Environment]::GetEnvironmentVariable(“PSModulePath”) + ";" + $DMEModulePath + "\OperationsManager\Powershell\"

[System.Environment]::SetEnvironmentVariable("PSModulePath", $NewPSPath)

Import-Module -Name OperationsManager


$env:USERPROFILE
