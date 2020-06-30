<#
.NAME
    Downloaded Files Sorter
.SYNOPSIS
  Moves files from a target folder to Videos, Pictures, Documents and Program Installers folders based on file extensions.
.DESCRIPTION
  In the target folder all files are checked and moved to Videos, Pictures, Documents and Program Installers folders based on file 
  extensions. Paths to target folder and to user libraries can be defined by the user, and saved permanently to local computer, but
  there are default values for those paths. File sorting is not performed in subfolders of the target folder and they are not moved
  by the script.
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
$documentExtensions = "*.DOC", "*.DOCX", "*.HTML", "*.HTM", "*.ODT", "*.PDF", "*.XLS", "*.XLSX", "*.ODS", "*.PPT", "*.PPTX", "*.TXT"
$installerExtensions = "*.exe", "*.msi", "*.msm", "*.msp", "*.mst", "*.idt", "*.idt", "*.cub", "*.pcp"
$pictureExtensions = "*.JPG", "*.PNG", "*.GIF", "*.WEBP", "*.TIFF", "*.SD", "*.RAW", "*.BMP", "*.HEIF", "*.INDD", "*.JPEG", "*.SVG", "*.AI", "*.EPS", "*.PDF"
$videoExtensions = "*.WEBM", "*.MPG", "*.MP2", "*.MPEG", "*.MPE", "*.MPV", "*.OGG", "*.MP4", "*.M4P", "*.M4V", "*.AVI", "*.WMV", "*.MO", "*.QT", "*.FLV", "*.SWF", "*.AVCHD"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Timestamp
$timestamp = Get-Date -Format "dd.MM/yyyy HH:mm:ss"

#Default Paths
$defaultTargetFolder = ################################################################env:USERPROFILE
$defaultPicturesFolder = [environment]::getfolderpath("mypictures")
$defaultProgramInstallersFolder = "D:\Program Installers\"
$defaultDocumentsFolder = [environment]::getfolderpath("mydocuments")
$defaultVideosFolder = [environment]::getfolderpath("myvideos")

#Current Paths
$targetFolder = $defaultTargetFolder
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
Function Create-Item
{
    param($item, $type)

    if((Test-Path $item) -eq $False)
    {
        New-Item -Path $item -ItemType $type
    }
}

Function Get-Timestamp
{
    $timestamp = Get-Date -Format "yyyy.MM.dd. HH:mm:ss"
}

Function Move-Files
{
    param($extensions, $destination)

    $logEntry = $timestamp + " - File sorting started"

    Log-Write -LogPath $logFile -LineValue $logEntry

    foreach($extension in $extensions)
    {

        $path = $sorce + $extension

        $file = Get-ChildItem -Path $path

        Move-Item -Path $file -Destination $destination

        $logEntry = $file.Name + " moved to " + $destination

        Log-Write -LogPath $logFile -LineValue $logEntry
    }
}

Function Sort-Files
{
    Begin
    {
        $logEntry = "File sorting started - " + Get-Date -Format h

        Log-Write -LogPath $logFile -LineValue $logEntry
    }
  
    Process
    {
        Try
        {
            Move-Files -Extensions $documentExtensions -Destination $documentsFolder
            Move-Files -Extensions $pictureExtensions -Destination $picturesFolder
            Move-Files -Extensions $videoExtensions -Destination $videosFolder
            Move-Files -Extensions $installerExtensions -Destination $documentsFolder
        }
    
        Catch
        {
            Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
            Break
        }
    }
  
    End
    {
        If($true)
        {
            Log-Write -LogPath $sLogFile -LineValue "Completed Successfully."
            Log-Write -LogPath $sLogFile -LineValue "================================================================================================="
        }
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Create-Item -Item $appPath -Type Directory
Create-Item -Item $$logFile -Type File
Create-Item -Item $costumFoldersFile -Type file

Log-Finish -LogPath $logFile