<#
.SYNOPSIS
Creates necessary files and folders for the application

.DESCRIPTION
Crates files and folders with parameterized path and type only if specified file does not already exist.

.PARAMETER Item
Full name of file or folder. If it is a file extension is included.

.PARAMETER Type
Item type (File, Directory)

.EXAMPLE
New-ItemConditionalCreation -Item "C:\Example.txt" -Type File

.EXAMPLE
New-ItemConditionalCreation "C:\Example.txt" "File"

.EXAMPLE
New-ItemConditionalCreation "C:\Example" "Directory"

.EXAMPLE
"C:\Example" | New-ItemConditionalCreation 'Directory'

.NOTES
Version:        1.1
Author:         Zoran Jankov
#>
function New-ItemConditionalCreation {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true,
                   Position = 0,
                   ValueFromPipeline = $false,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "Full name of file or folder. If it is a file extension is included.")]
        [string]
        $Item,

        [Parameter(Mandatory = $true,
                   Position = 1,
                   ValueFromPipeline = $false,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "Item type (File, Directory)")]
        [string]
        $Type
    )

    process {
        if ((Test-Path $Item) -eq $false) {
            New-Item -Path $Item -ItemType $Type
            Write-Log -Message "Successfully created $Item $Type"
        }
    }
}