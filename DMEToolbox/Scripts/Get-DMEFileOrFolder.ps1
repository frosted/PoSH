<#
.Synopsis
Prompt user to select file or folder
.DESCRIPTION
This function can be leveraged in other scripts when a path to a file or folder is needed (e.g path to software repository)

.PARAMETER InitialDirectory
Directory path for file/folder dialogue to start in

.PARAMETER Type
Whether you're looking for a File or Folder

.EXAMPLE
# prompt user for file
Get-DMEFileOrFolder -Type File

.EXAMPLE
# prompt user for folder, starting in c:\temp
Get-DMEFileOrFolder -Type Folder -InitialDirectory c:\temp

.NOTES
Author: Ed Frost
Date:   2023.06.15
#>
Function Get-DMEFileOrFolder
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $False,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'Parameter Set 1')]
        [string]
        $InitialDirectory = [Environment]::GetFolderPath('Desktop'),
        [ValidateSet('File', 'Folder')]
        [string]
        $Type = 'File'
    )

    Begin
    {
        Add-Type -AssemblyName System.Windows.Forms
    }
    Process
    {
        If (-not(Test-Path -Path $InitialDirectory -ErrorAction SilentlyContinue)) 
        { 
            $InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        }

        Switch ($Type)
        {
            'File'
            {
                $FSO = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $InitialDirectory }
            }
            'Folder'
            {
                $FSO = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ SelectedPath = $InitialDirectory }
            }
        }
        $null = $FSO.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true; TopLevel = $true }))
    }
    End
    {
        Switch ($Type)
        {
            'File'
            {
                Return Get-Item -Path $FSO.FileName | Select-Object *
            }
            'Folder'
            {
                Return Get-Item -Path $FSO.SelectedPath | Select-Object *
            }
        }
    }
}