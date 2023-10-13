# Universal psm file
# Requires -Version 5.0

# Get functions files

$Functions = @( Get-ChildItem -Path $PSScriptRoot\Scripts -Filter *.ps1 -ErrorAction SilentlyContinue)

# Dot source the files
Foreach ($Function in $Functions) {
    try {
        . $Function.Fullname
    } catch {
        Write-Error -Message "Failed to import function $($Function.fullname)"
    }
}

# Export everything in public folder
Export-ModuleMember -Function * -Cmdlet * -Alias *
Initialize-DMEEnvironment "$(Split-Path $MyInvocation.MyCommand.Path)\DMEEnvironment.conf" 
New-PSDrive -Root (Get-ChildItem $DMEConfigPath).DirectoryName -Name DME -PSProvider FileSystem -Scope global -ErrorAction SilentlyContinue

Write-Host "DME" -ForegroundColor RED -NoNewline
Write-Host "Toolbox Module" 
Set-Location -Path DME: