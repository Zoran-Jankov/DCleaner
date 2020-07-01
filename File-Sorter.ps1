<#
.NAME
    Downloaded Files Sorter
.SYNOPSIS
  Moves files from a source folder to Videos, Pictures, Documents and Program Installers folders based on file extensions.
.DESCRIPTION
  In the source folder all files are checked and moved to Videos, Pictures, Documents and Program Installers folders based on file 
  extensions. Paths to source folder and to user libraries can be defined by the user, and can besaved permanently to local 
  computer, butthere are default values for those paths. File sorting is not performed in subfolders of the target folder and they 
  are not moved by the script.
.INPUTS
  Path to target, Videos, Documents, Pictures and Program Installers folders can by the user, and there are tree buttons:
  Default Locations - restores default paths for all folders.
  Save Locations - save user defined paths to folders to local file.
  Sort Files - executes the script for file sorting.
.OUTPUTS
  Log file stored in "%APPDATA%\File-Sorter\File-Sorter-Log.log"
.NOTES
  Version:        0.1
  Author:         Zoran Jankov
  Creation Date:  30.06.2020.
  Purpose/Change: Initial script development
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Set known extensions
$documentExtensions = "*.DOC", "*.DOCX", "*.HTML", "*.HTM", "*.ODT", "*.PDF", "*.XLS", "*.XLSX", "*.ODS", "*.PPT", "*.PPTX", "*.TXT", "*.LOG"
$installerExtensions = "*.exe", "*.msi", "*.msm", "*.msp", "*.mst", "*.idt", "*.idt", "*.cub", "*.pcp"
$pictureExtensions = "*.JPG", "*.PNG", "*.GIF", "*.WEBP", "*.TIFF", "*.SD", "*.RAW", "*.BMP", "*.HEIF", "*.INDD", "*.JPEG", "*.SVG", "*.AI", "*.EPS", "*.PDF"
$videoExtensions = "*.WEBM", "*.MPG", "*.MP2", "*.MPEG", "*.MPE", "*.MPV", "*.OGG", "*.MP4", "*.M4P", "*.M4V", "*.AVI", "*.WMV", "*.MO", "*.QT", "*.FLV", "*.SWF", "*.AVCHD"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Timestamp Definition
$timestamp = Get-Date -Format "dd.MM/yyyy HH:mm:ss"

#Default Paths
$defaultSourceFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$defaultPicturesFolder = [environment]::getfolderpath("mypictures")
$defaultProgramInstallersFolder = "D:\Program Installers\"
$defaultDocumentsFolder = [environment]::getfolderpath("mydocuments")
$defaultVideosFolder = [environment]::getfolderpath("myvideos")

#Current Paths
$sourceFolder = $defaultTargetFolder
$picturesFolder = $defaultPicturesFolder
$programInstallersFolder = $defaultProgramInstallersFolder 
$documentsFolder = $defaultDocumentsFolder
$videosFolder = $defaultVideosFolder

#Aplication Folder Info
$appPath = $env:APPDATA + "\File-Sorter"

#Log File Info
$logName = "File-Sorter-Log.log"
$logFile = Join-Path -Path $appPath -ChildPath $logName

#Costum Folder Location File Info
$costumFoldersName = "Costum-Folders.cvs"
$costumFoldersFile = Join-Path -Path $appPath -ChildPath $costumFoldersName 

#-----------------------------------------------------------[Functions]------------------------------------------------------------

#Creates necessary files and directory in %APPDATA% directory
Function New-ItemConditionalCreation
{
    param($Item, $Type)

    if((Test-Path $Item) -eq $False)
    {
        New-Item -Path $Item -ItemType $Type
    }
}

#Refreshes formatted timestamp variable - $timestamp
Function Get-Timestamp
{
    $timestamp = Get-Date -Format "yyyy.MM.dd. HH:mm:ss"
}

#Moves files with defined extensions from source folder to defined destination folder
Function Move-Files
{
    param($Extensions, $Destination)

    Get-Timestamp

    $logEntry = $timestamp + " - Started moving files to " + $Destination

    Log-Write -LogPath $logFile -LineValue $logEntry

    Try
    {
        foreach($extension in $Extensions)
        {
            $path = $sourceFolder + $extension

            $file = Get-ChildItem -Path $path

            Move-Item -Path $file -Destination $Destination

            Get-Timestamp

            $logEntry = $timestamp + " - " + $file.Name + " moved to " + $Destination

            Log-Write -LogPath $logFile -LineValue $logEntry
        }
    }

    Catch
    {
        Log-Error -LogPath $logFile -ErrorDesc $_.Exception -ExitGracefully $True
        Break
    }
}

#Starts files moving from source to user library folders
Function Start-FileSorting
{
    Get-Timestamp

    $logEntry = $timestamp + " - File sorting started"

    Log-Write -LogPath $logFile -LineValue $logEntry
   
    Move-Files -Extensions $documentExtensions -Destination $documentsFolder
    Move-Files -Extensions $pictureExtensions -Destination $picturesFolder
    Move-Files -Extensions $videoExtensions -Destination $videosFolder
    Move-Files -Extensions $installerExtensions -Destination $programInstallersFolder

    End
    {
        If($true)
        {
            Log-Write -LogPath $logFile -LineValue "Completed Successfully."
            Log-Write -LogPath $logFile -LineValue "=============================================================================="
        }
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Create-Item -Item $appPath -Type Directory
Create-Item -Item $logFile -Type File

Create-Item -Item $costumFoldersFile -Type file

Log-Finish -LogPath $logFile