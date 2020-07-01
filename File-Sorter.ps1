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
  Version:        0.9
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

#Create GUI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$formFileSorter                  = New-Object system.Windows.Forms.Form
$formFileSorter.ClientSize       = New-Object System.Drawing.Point(795,245)
$formFileSorter.text             = "File Sorter"
$formFileSorter.TopMost          = $true

$lblDownloads                    = New-Object system.Windows.Forms.Label
$lblDownloads.text               = "Downloads"
$lblDownloads.AutoSize           = $true
$lblDownloads.width              = 25
$lblDownloads.height             = 10
$lblDownloads.location           = New-Object System.Drawing.Point(20,20)
$lblDownloads.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtDownloads                    = New-Object system.Windows.Forms.TextBox
$txtDownloads.multiline          = $false
$txtDownloads.width              = 600
$txtDownloads.height             = 20
$txtDownloads.location           = New-Object System.Drawing.Point(170,16)
$txtDownloads.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$txtDownloads.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#9b9b9b")

$lblDocuments                    = New-Object system.Windows.Forms.Label
$lblDocuments.text               = "Documents"
$lblDocuments.AutoSize           = $true
$lblDocuments.width              = 25
$lblDocuments.height             = 10
$lblDocuments.location           = New-Object System.Drawing.Point(20,50)
$lblDocuments.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtDocuments                    = New-Object system.Windows.Forms.TextBox
$txtDocuments.multiline          = $false
$txtDocuments.width              = 600
$txtDocuments.height             = 20
$txtDocuments.location           = New-Object System.Drawing.Point(170,46)
$txtDocuments.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblPictures                     = New-Object system.Windows.Forms.Label
$lblPictures.text                = "Pictures"
$lblPictures.AutoSize            = $true
$lblPictures.width               = 25
$lblPictures.height              = 10
$lblPictures.location            = New-Object System.Drawing.Point(20,80)
$lblPictures.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtPictures                     = New-Object system.Windows.Forms.TextBox
$txtPictures.multiline           = $false
$txtPictures.width               = 600
$txtPictures.height              = 20
$txtPictures.location            = New-Object System.Drawing.Point(170,76)
$txtPictures.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblVideos                       = New-Object system.Windows.Forms.Label
$lblVideos.text                  = "Videos"
$lblVideos.AutoSize              = $true
$lblVideos.width                 = 25
$lblVideos.height                = 10
$lblVideos.location              = New-Object System.Drawing.Point(20,110)
$lblVideos.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtVideos                       = New-Object system.Windows.Forms.TextBox
$txtVideos.multiline             = $false
$txtVideos.width                 = 600
$txtVideos.height                = 20
$txtVideos.location              = New-Object System.Drawing.Point(170,107)
$txtVideos.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnSortFiles                    = New-Object system.Windows.Forms.Button
$btnSortFiles.text               = "Sort Files"
$btnSortFiles.width              = 150
$btnSortFiles.height             = 30
$btnSortFiles.location           = New-Object System.Drawing.Point(325,200)
$btnSortFiles.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$lblProgramInstallers            = New-Object system.Windows.Forms.Label
$lblProgramInstallers.text       = "Program Installers"
$lblProgramInstallers.AutoSize   = $true
$lblProgramInstallers.width      = 25
$lblProgramInstallers.height     = 10
$lblProgramInstallers.location   = New-Object System.Drawing.Point(20,140)
$lblProgramInstallers.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 0
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(144,140)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtProgramInstallers            = New-Object system.Windows.Forms.TextBox
$txtProgramInstallers.multiline  = $false
$txtProgramInstallers.width      = 600
$txtProgramInstallers.height     = 20
$txtProgramInstallers.location   = New-Object System.Drawing.Point(170,136)
$txtProgramInstallers.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnDefaultLocations             = New-Object system.Windows.Forms.Button
$btnDefaultLocations.text        = "Default Locations"
$btnDefaultLocations.width       = 150
$btnDefaultLocations.height      = 30
$btnDefaultLocations.location    = New-Object System.Drawing.Point(50,200)
$btnDefaultLocations.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnSaveLocations                = New-Object system.Windows.Forms.Button
$btnSaveLocations.text           = "Save Locations"
$btnSaveLocations.width          = 150
$btnSaveLocations.height         = 30
$btnSaveLocations.location       = New-Object System.Drawing.Point(590,200)
$btnSaveLocations.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$formFileSorter.controls.AddRange(@($lblDownloads,$txtDownloads,$lblDocuments,$txtDocuments,$lblPictures,$txtPictures,$lblVideos,
$txtVideos,$btnSortFiles,$lblProgramInstallers,$TextBox1,$txtProgramInstallers,$btnDefaultLocations,$btnSaveLocations))

$btnSortFiles.Add_Click({ Start-FileSorting })
$btnDefaultLocations.Add_Click({ Set-DefaultLocations })
$btnSaveLocations.Add_Click({ Save-FolderSettings })

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

#Load Locations Folders
if(Test-Path costumFoldersFile -eq $True)
{
    
}

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
#Resets folder locations to default values
Set-DefaultLocations
{
    $txtDownloads.Text = [string]$defaultTargetFolder
    $txtPictures.Text = [string]$defaultPicturesFolder
    $txtProgramInstallers.Text = [string]$defaultProgramInstallersFolder 
    $txtDocuments.Text = [string]$defaultDocumentsFolder
    $txtVideos.Text = [string]$defaultVideosFolder
}

#Saves costum folder locations to local file
Function Save-FolderSettings
{
    if((Test-Path $costumFoldersFile) -eq $False)
    {
        New-Item -Path $costumFoldersFile -ItemType File
        Add-Content -Path $costumFoldersFile -Value '"FolderName","Path"'
    }
    else
    {
        Import-Csv $costumFoldersFile
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
        If($True)
        {
            Log-Write -LogPath $logFile -LineValue "Completed Successfully."
            Log-Write -LogPath $logFile -LineValue "=============================================================================="
        }
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Create-Item -Item $appPath -Type Directory
Create-Item -Item $logFile -Type File

[void]$formFileSorter.ShowDialog()

Log-Finish -LogPath $logFile