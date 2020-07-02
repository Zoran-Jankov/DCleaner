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
  Version:        0.14
  Author:         Zoran Jankov
  Creation Date:  30.06.2020.
  Purpose/Change: Initial script development
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Set known extensions
$global:documentExtensions = "*.DOC", "*.DOCX", "*.HTML", "*.HTM", "*.ODT", "*.PDF", "*.XLS", "*.XLSX", "*.ODS", "*.PPT", "*.PPTX", "*.TXT", "*.LOG"
$global:installerExtensions = "*.exe", "*.msi", "*.msm", "*.msp", "*.mst", "*.idt", "*.idt", "*.cub", "*.pcp"
$global:pictureExtensions = "*.JPG", "*.PNG", "*.GIF", "*.WEBP", "*.TIFF", "*.SD", "*.RAW", "*.BMP", "*.HEIF", "*.INDD", "*.JPEG", "*.SVG", "*.AI", "*.EPS", "*.PDF"
$global:videoExtensions = "*.WEBM", "*.MPG", "*.MP2", "*.MPEG", "*.MPE", "*.MPV", "*.OGG", "*.MP4", "*.M4P", "*.M4V", "*.AVI", "*.WMV", "*.MO", "*.QT", "*.FLV", "*.SWF", "*.AVCHD"

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

$global:txtDownloads                    = New-Object system.Windows.Forms.TextBox
$global:txtDownloads.multiline          = $false
$global:txtDownloads.width              = 600
$global:txtDownloads.height             = 20
$global:txtDownloads.location           = New-Object System.Drawing.Point(170,16)
$global:txtDownloads.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblDocuments                    = New-Object system.Windows.Forms.Label
$lblDocuments.text               = "Documents"
$lblDocuments.AutoSize           = $true
$lblDocuments.width              = 25
$lblDocuments.height             = 10
$lblDocuments.location           = New-Object System.Drawing.Point(20,50)
$lblDocuments.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$global:txtDocuments                    = New-Object system.Windows.Forms.TextBox
$global:txtDocuments.multiline          = $false
$global:txtDocuments.width              = 600
$global:txtDocuments.height             = 20
$global:txtDocuments.location           = New-Object System.Drawing.Point(170,46)
$global:txtDocuments.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblPictures                     = New-Object system.Windows.Forms.Label
$lblPictures.text                = "Pictures"
$lblPictures.AutoSize            = $true
$lblPictures.width               = 25
$lblPictures.height              = 10
$lblPictures.location            = New-Object System.Drawing.Point(20,80)
$lblPictures.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$global:txtPictures                     = New-Object system.Windows.Forms.TextBox
$global:txtPictures.multiline           = $false
$global:txtPictures.width               = 600
$global:txtPictures.height              = 20
$global:txtPictures.location            = New-Object System.Drawing.Point(170,76)
$global:txtPictures.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblVideos                       = New-Object system.Windows.Forms.Label
$lblVideos.text                  = "Videos"
$lblVideos.AutoSize              = $true
$lblVideos.width                 = 25
$lblVideos.height                = 10
$lblVideos.location              = New-Object System.Drawing.Point(20,110)
$lblVideos.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$global:txtVideos                       = New-Object system.Windows.Forms.TextBox
$global:txtVideos.multiline             = $false
$global:txtVideos.width                 = 600
$global:txtVideos.height                = 20
$global:txtVideos.location              = New-Object System.Drawing.Point(170,107)
$global:txtVideos.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnSortFiles                    = New-Object system.Windows.Forms.Button
$btnSortFiles.text               = "Sort Files"
$btnSortFiles.width              = 150
$btnSortFiles.height             = 30
$btnSortFiles.location           = New-Object System.Drawing.Point(325,200)
$btnSortFiles.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10,
                                   [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$lblProgramInstallers            = New-Object system.Windows.Forms.Label
$lblProgramInstallers.text       = "Program Installers"
$lblProgramInstallers.AutoSize   = $true
$lblProgramInstallers.width      = 25
$lblProgramInstallers.height     = 10
$lblProgramInstallers.location   = New-Object System.Drawing.Point(20,140)
$lblProgramInstallers.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$global:txtProgramInstallers            = New-Object system.Windows.Forms.TextBox
$global:txtProgramInstallers.multiline  = $false
$global:txtProgramInstallers.width      = 600
$global:txtProgramInstallers.height     = 20
$global:txtProgramInstallers.location   = New-Object System.Drawing.Point(170,136)
$global:txtProgramInstallers.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

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

$formFileSorter.controls.AddRange(@($lblDownloads,$global:txtDownloads,$lblDocuments,$global:txtDocuments,$lblPictures,
$global:txtPictures,$lblVideos,$global:txtVideos,$btnSortFiles,$lblProgramInstallers,$global:txtProgramInstallers,
$btnDefaultLocations,$btnSaveLocations))

$btnSortFiles.Add_Click({ Start-FileSorting })
$btnDefaultLocations.Add_Click({ Set-DefaultLocations })
$btnSaveLocations.Add_Click({ Save-FolderSettings })

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Default Paths
$global:defaultSourceFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$global:defaultPicturesFolder = [environment]::getfolderpath("mypictures")
$global:defaultProgramInstallersFolder = "D:\Program Installers\"
$global:defaultDocumentsFolder = [environment]::getfolderpath("mydocuments")
$global:defaultVideosFolder = [environment]::getfolderpath("myvideos")

#Aplication Folder Info
$global:appPath = $env:APPDATA + "\File-Sorter"

#Log File Info
$global:logName = "File-Sorter-Log.log"
$global:logFile = Join-Path -Path $global:appPath -ChildPath $global:logName

#Costum Folder Location File Info
$global:costumFoldersName = "Costum-Folders.cvs"
$global:costumFoldersFile = Join-Path -Path $global:appPath -ChildPath $global:costumFoldersName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

#Creates necessary files and folders in %APPDATA% folder
Function New-ItemConditionalCreation
{
    param($Item, $Type)

    if((Test-Path $Item) -eq $False)
    {
        New-Item -Path $Item -ItemType $Type
    }
}

#Writes log entry
Function Write-Log
{
    param($Message)

    $timestamp = Get-Date -Format "yyyy.MM.dd. HH:mm:ss"

    $logEntry = $timestamp + " - " + $Message 

    Log-Write -LogPath $global:logFile -LineValue $logEntry
}

#Moves files with defined extensions from source folder to defined destination folder
Function Move-Files
{
    param($Extensions, $Destination)

    $massage = "Started moving files to " + $Destination

    Write-Log -Message $massage

    Try
    {
        foreach($extension in $Extensions)
        {
            $path = $global:sourceFolder + $extension

            Get-ChildItem -Path $path | Move-Item -Destination $Destination

        }
    }

    Catch
    {
        Log-Error -LogPath $global:logFile -ErrorDesc $_.Exception -ExitGracefully $True
        Break
    }
}

#Starts files moving from source to user library folders
Function Start-FileSorting
{
    Write-Log -Message "File sorting started"

    #Get Locations
    $global:sourceFolder = $global:txtDownloads.Text
    $documentsFolder = $global:txtDocuments.Text
    $picturesFolder = $global:txtPictures.Text
    $videosFolder = $global:txtVideos.Text
    $programInstallersFolder = $global:txtProgramInstallers.Text

    #Moves files from source to user library folders 
    Move-Files -Extensions $globaldocumentExtensions -Destination $documentsFolder
    Move-Files -Extensions $globalpictureExtensions -Destination $picturesFolder
    Move-Files -Extensions $globalvideoExtensions -Destination $videosFolder
    Move-Files -Extensions $globalinstallerExtensions -Destination $programInstallersFolder
    
    Log-Write -LogPath $global:logFile -LineValue "Completed Successfully."
    Log-Write -LogPath $global:logFile -LineValue "==============================================================================="
}

#Resets folder locations to default values
Function Set-DefaultLocations
{
    $global:txtDownloads.Text = $global:defaultSourceFolder
    $global:txtPictures.Text = $global:defaultPicturesFolder
    $global:txtProgramInstallers.Text = $global:defaultProgramInstallersFolder
    $global:txtDocuments.Text = $global:defaultDocumentsFolder
    $global:txtVideos.Text = $global:defaultVideosFolder
}

#Saves costum folder locations to local file
Function Save-FolderSettings
{
    if((Test-Path $global:costumFoldersFile) -eq $False)
    {
        New-Item -Path $global:costumFoldersFile -ItemType File
        Add-Content -Path $global:costumFoldersFile -Value '"FolderName";"Path"'
    }
    else
    {
        Import-Csv $global:costumFoldersFile
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Load Locations Folders
if((Test-Path $global:costumFoldersFile) -eq $True)
{
    Set-DefaultLocations
}

New-ItemConditionalCreation -Item $global:appPath -Type Directory
New-ItemConditionalCreation -Item $global:logFile -Type File

[void]$formFileSorter.ShowDialog()