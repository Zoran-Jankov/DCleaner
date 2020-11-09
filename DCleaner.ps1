<#
.SYNOPSIS
Moves files from a source folder to Music, Videos, Pictures, Documents and Program Installers folders based on file extensions.

.DESCRIPTION
This application moves files from the source folder to user libraries. In the source folder all files are checked and moved to
Videos, Pictures, Documents and Program Installers folders based on file extensions. Paths to target folder and to user libraries
can be defined by the user, and saved permanently to local computer, but there are default values for those paths.
File sorting is not performed in subfolders of the target folder and they are not moved by the script.

.INPUTS
Path to target, Videos, Documents, Pictures and Program Installers folders can by the user, and there are tree buttons:
Default Locations - restores default paths for all folders.
Save Locations - save user defined paths to folders to local file.
Sort Files - executes the script for file sorting.

.OUTPUTS
Log file stored in "%APPDATA%\File Sorter\File-Sorter-Log.log"
Custom folder locations are saved to "%APPDATA%\File Sorter\Custom-Folders.xml"

.NOTES
    Version:        1.0
    Author:         Zoran Jankov
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Set known extensions
$documentExtensions = "*.DOC", "*.DOCX", "*.HTML", "*.HTM", "*.ODT", "*.PDF", "*.XLS", "*.XLSX", "*.ODS", "*.PPT", "*.PPTX", "*.TXT", "*.LOG"
$installerExtensions = "*.exe", "*.msi", "*.msm", "*.msp", "*.mst", "*.idt", "*.idt", "*.cub", "*.pcp", "*.jar"
$pictureExtensions = "*.JPG", "*.PNG", "*.GIF", "*.WEBP", "*.TIFF", "*.SD", "*.RAW", "*.BMP", "*.HEIF", "*.INDD", "*.JPEG", "*.SVG", "*.AI", "*.EPS", "*.PDF", "*.cvs"
$videoExtensions = "*.WEBM", "*.MPG", "*.MP2", "*.MPEG", "*.MPE", "*.MPV", "*.OGG", "*.MP4", "*.M4P", "*.M4V", "*.AVI", "*.WMV", "*.MO", "*.QT", "*.FLV", "*.SWF", "*.AVCHD"

#Default Paths
$defaultSourceFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$defaultPicturesFolder = [environment]::getfolderpath("mypictures")
$defaultProgramInstallersFolder = "D:\Program Installers\"
$defaultDocumentsFolder = [environment]::getfolderpath("mydocuments")
$defaultVideosFolder = [environment]::getfolderpath("myvideos")

#Aplication Folder Info
$AppPath = "$Env:APPDATA\DCleaner"

#Log File Info
$LogName = "$AppPath\DCleaner-Log.log"

#Custom Folder Location File Info
$customFoldersName = "$AppPath\Custom-Folders.csv"

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Import-Module "$PSScriptRoot\Modules\New-ItemConditionalCreation.psm1"

#Moves files with defined extensions from source folder to defined destination folder
function Move-Files
{
    param([String]$Extensions, [String]$Source, [String]$Destination)

    $massage = "Started moving files to " + $Destination
    Write-Log -Message $massage

    Try
    {
        foreach($extension in $Extensions)
        {
            $path = Join-Path -Path $Source -ChildPath $extension
            $files = Get-ChildItem -Path $path

            foreach($file in $files)
            {
                Move-Item -Path $file.FullName -Destination $Destination
                $logEntry = $file.Name + " moved to " + $Destination
                Write-Log -Message $logEntry
            }
        }
    }

    Catch
    {
        Write-Log -Message $_.Exception
        Break
    }

    $massage = "Finished moving files to " + $Destination
    Write-Log -Message $massage
}

#Starts files moving from source to user library folders
function Start-FileSorting
{
    Write-Log -Message "File sorting started"

    #Get Locations
    $sourceFolder = $txtDownloads.Text
    $documentsFolder = $txtDocuments.Text
    $picturesFolder = $txtPictures.Text
    $videosFolder = $txtVideos.Text
    $programInstallersFolder = $txtProgramInstallers.Text

    #Moves files from source to user library folders
    Move-Files -Extensions $documentExtensions -Source $sourceFolder -Destination $documentsFolder
    Move-Files -Extensions $pictureExtensions -Source $sourceFolder -Destination $picturesFolder
    Move-Files -Extensions $videoExtensions -Source $sourceFolder -Destination $videosFolder
    Move-Files -Extensions $installerExtensions -Source $sourceFolder -Destination $programInstallersFolder

    Write-Log -Message "Completed Successfully."
    Add-Content -Path $logFile -Value "==========================================================================================="
}

#Resets folder locations to default values
function Set-DefaultLocations
{
    $txtDownloads.Text = $defaultSourceFolder
    $txtPictures.Text = $defaultPicturesFolder
    $txtProgramInstallers.Text = $defaultProgramInstallersFolder
    $txtDocuments.Text = $defaultDocumentsFolder
    $txtVideos.Text = $defaultVideosFolder
}

#Saves custom folder locations to local file
function Save-FolderSettings
{
        #TODO Save paths
}

#Save folder path
function Save-Path
{
    param ([string]$Name, [string]$Path)
    
    $xmlWriter.WriteStartElement($Name)

    $xmlWriter.WriteElementString("Path",$Path)

    $xmlWriter.WriteEndElement()
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Create GUI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$DCleanerForm                      = New-Object system.Windows.Forms.Form
$DCleanerForm.ClientSize           = New-Object System.Drawing.Point(741,383)
$DCleanerForm.text                 = "DCleaner"
$DCleanerForm.TopMost              = $true

$SourceFolderLabel                 = New-Object system.Windows.Forms.Label
$SourceFolderLabel.text            = "Source Folder"
$SourceFolderLabel.AutoSize        = $true
$SourceFolderLabel.width           = 25
$SourceFolderLabel.height          = 10
$SourceFolderLabel.location        = New-Object System.Drawing.Point(325,14)
$SourceFolderLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$DestinationFoldersLabel           = New-Object system.Windows.Forms.Label
$DestinationFoldersLabel.text      = "Destination Folders"
$DestinationFoldersLabel.AutoSize  = $true
$DestinationFoldersLabel.width     = 25
$DestinationFoldersLabel.height    = 10
$DestinationFoldersLabel.location  = New-Object System.Drawing.Point(300,90)
$DestinationFoldersLabel.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$DownloadsFolderLabel              = New-Object system.Windows.Forms.Label
$DownloadsFolderLabel.text         = "Downloads Folder"
$DownloadsFolderLabel.AutoSize     = $true
$DownloadsFolderLabel.width        = 25
$DownloadsFolderLabel.height       = 10
$DownloadsFolderLabel.location     = New-Object System.Drawing.Point(25,53)
$DownloadsFolderLabel.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$DownloadsFolderTextBox            = New-Object system.Windows.Forms.TextBox
$DownloadsFolderTextBox.multiline  = $false
$DownloadsFolderTextBox.width      = 460
$DownloadsFolderTextBox.height     = 20
$DownloadsFolderTextBox.enabled    = $false
$DownloadsFolderTextBox.location   = New-Object System.Drawing.Point(180,49)
$DownloadsFolderTextBox.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$DownloadsFolderButton             = New-Object system.Windows.Forms.Button
$DownloadsFolderButton.text        = "..."
$DownloadsFolderButton.width       = 45
$DownloadsFolderButton.height      = 27
$DownloadsFolderButton.location    = New-Object System.Drawing.Point(662,47)
$DownloadsFolderButton.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$DocumentsFolderLabel              = New-Object system.Windows.Forms.Label
$DocumentsFolderLabel.text         = "Documents Folder"
$DocumentsFolderLabel.AutoSize     = $true
$DocumentsFolderLabel.width        = 25
$DocumentsFolderLabel.height       = 10
$DocumentsFolderLabel.location     = New-Object System.Drawing.Point(25,130)
$DocumentsFolderLabel.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$DocumentFolderTextBox             = New-Object system.Windows.Forms.TextBox
$DocumentFolderTextBox.multiline   = $false
$DocumentFolderTextBox.width       = 460
$DocumentFolderTextBox.height      = 20
$DocumentFolderTextBox.enabled     = $false
$DocumentFolderTextBox.location    = New-Object System.Drawing.Point(180,127)
$DocumentFolderTextBox.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$DocumentsFolderButton             = New-Object system.Windows.Forms.Button
$DocumentsFolderButton.text        = "..."
$DocumentsFolderButton.width       = 45
$DocumentsFolderButton.height      = 27
$DocumentsFolderButton.location    = New-Object System.Drawing.Point(662,124)
$DocumentsFolderButton.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$MusicFolderLabel                  = New-Object system.Windows.Forms.Label
$MusicFolderLabel.text             = "Music Folder"
$MusicFolderLabel.AutoSize         = $true
$MusicFolderLabel.width            = 25
$MusicFolderLabel.height           = 10
$MusicFolderLabel.location         = New-Object System.Drawing.Point(25,166)
$MusicFolderLabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$MusicFolderTextBox                = New-Object system.Windows.Forms.TextBox
$MusicFolderTextBox.multiline      = $false
$MusicFolderTextBox.width          = 460
$MusicFolderTextBox.height         = 20
$MusicFolderTextBox.enabled        = $false
$MusicFolderTextBox.location       = New-Object System.Drawing.Point(180,163)
$MusicFolderTextBox.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$MusicFolderButton                 = New-Object system.Windows.Forms.Button
$MusicFolderButton.text            = "..."
$MusicFolderButton.width           = 45
$MusicFolderButton.height          = 27
$MusicFolderButton.location        = New-Object System.Drawing.Point(662,161)
$MusicFolderButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$PicturesFolderTextBox             = New-Object system.Windows.Forms.TextBox
$PicturesFolderTextBox.multiline   = $false
$PicturesFolderTextBox.width       = 460
$PicturesFolderTextBox.height      = 20
$PicturesFolderTextBox.enabled     = $false
$PicturesFolderTextBox.location    = New-Object System.Drawing.Point(180,198)
$PicturesFolderTextBox.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$PicturesFolderLabel               = New-Object system.Windows.Forms.Label
$PicturesFolderLabel.text          = "Pictures Folder"
$PicturesFolderLabel.AutoSize      = $true
$PicturesFolderLabel.width         = 25
$PicturesFolderLabel.height        = 10
$PicturesFolderLabel.location      = New-Object System.Drawing.Point(25,201)
$PicturesFolderLabel.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$PicturesFolderButton              = New-Object system.Windows.Forms.Button
$PicturesFolderButton.text         = "..."
$PicturesFolderButton.width        = 45
$PicturesFolderButton.height       = 27
$PicturesFolderButton.location     = New-Object System.Drawing.Point(662,195)
$PicturesFolderButton.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$VideosFolderTextBox               = New-Object system.Windows.Forms.TextBox
$VideosFolderTextBox.multiline     = $false
$VideosFolderTextBox.width         = 460
$VideosFolderTextBox.height        = 20
$VideosFolderTextBox.enabled       = $false
$VideosFolderTextBox.location      = New-Object System.Drawing.Point(180,234)
$VideosFolderTextBox.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$VideosFolderLabel                 = New-Object system.Windows.Forms.Label
$VideosFolderLabel.text            = "Videos Folder"
$VideosFolderLabel.AutoSize        = $true
$VideosFolderLabel.width           = 25
$VideosFolderLabel.height          = 10
$VideosFolderLabel.location        = New-Object System.Drawing.Point(25,237)
$VideosFolderLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$VideoFolderButton                 = New-Object system.Windows.Forms.Button
$VideoFolderButton.text            = "..."
$VideoFolderButton.width           = 45
$VideoFolderButton.height          = 27
$VideoFolderButton.location        = New-Object System.Drawing.Point(662,232)
$VideoFolderButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$InstallersFolderTextBox           = New-Object system.Windows.Forms.TextBox
$InstallersFolderTextBox.multiline = $false
$InstallersFolderTextBox.width     = 460
$InstallersFolderTextBox.height    = 20
$InstallersFolderTextBox.enabled   = $false
$InstallersFolderTextBox.location  = New-Object System.Drawing.Point(180,269)
$InstallersFolderTextBox.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$InstallersFolderLabel             = New-Object system.Windows.Forms.Label
$InstallersFolderLabel.text        = "Installers Folder"
$InstallersFolderLabel.AutoSize    = $true
$InstallersFolderLabel.width       = 25
$InstallersFolderLabel.height      = 10
$InstallersFolderLabel.location    = New-Object System.Drawing.Point(25,273)
$InstallersFolderLabel.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$InstallersFolderButton            = New-Object system.Windows.Forms.Button
$InstallersFolderButton.text       = "..."
$InstallersFolderButton.width      = 45
$InstallersFolderButton.height     = 27
$InstallersFolderButton.location   = New-Object System.Drawing.Point(662,267)
$InstallersFolderButton.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$CleanButton                       = New-Object system.Windows.Forms.Button
$CleanButton.text                  = "Clean"
$CleanButton.width                 = 140
$CleanButton.height                = 30
$CleanButton.location              = New-Object System.Drawing.Point(213,330)
$CleanButton.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$DefaultLocationsButton            = New-Object system.Windows.Forms.Button
$DefaultLocationsButton.text       = "Default Locations"
$DefaultLocationsButton.width      = 140
$DefaultLocationsButton.height     = 30
$DefaultLocationsButton.location   = New-Object System.Drawing.Point(35,330)
$DefaultLocationsButton.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SaveLocationsButton               = New-Object system.Windows.Forms.Button
$SaveLocationsButton.text          = "Save Locations"
$SaveLocationsButton.width         = 140
$SaveLocationsButton.height        = 30
$SaveLocationsButton.location      = New-Object System.Drawing.Point(390,330)
$SaveLocationsButton.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$AboutButton                       = New-Object system.Windows.Forms.Button
$AboutButton.text                  = "About DCleaner"
$AboutButton.width                 = 140
$AboutButton.height                = 30
$AboutButton.location              = New-Object System.Drawing.Point(570,330)
$AboutButton.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$DCleanerForm.controls.AddRange(@($SourceFolderLabel,
                                  $DestinationFoldersLabel,
                                  $DownloadsFolderLabel,
                                  $DownloadsFolderTextBox,
                                  $DownloadsFolderButton,
                                  $DocumentsFolderLabel,
                                  $DocumentFolderTextBox,
                                  $DocumentsFolderButton,
                                  $MusicFolderLabel,
                                  $MusicFolderTextBox,
                                  $MusicFolderButton,
                                  $PicturesFolderLabel,
                                  $PicturesFolderTextBox,
                                  $PicturesFolderButton,
                                  $VideosFolderLabel,
                                  $VideosFolderTextBox,
                                  $VideoFolderButton,
                                  $InstallersFolderLabel,
                                  $InstallersFolderTextBox,
                                  $InstallersFolderButton,
                                  $CleanButton,
                                  $DefaultLocationsButton,
                                  $SaveLocationsButton,
                                  $AboutButton))

$DownloadsFolderButton.Add_Click({  })
$DocumentsFolderButton.Add_Click({  })
$MusicFolderButton.Add_Click({  })
$PicturesFolderButton.Add_Click({  })
$VideoFolderButton.Add_Click({  })
$InstallersFolderButton.Add_Click({  })
$DefaultLocationsButton.Add_Click({  })
$CleanButton.Add_Click({  })
$SaveLocationsButton.Add_Click({  })
$AboutButton.Add_Click({
    Start-Process https://github.com/Zoran-Jankov/DCleaner/blob/master/README.md
  })

[void]$DCleanerForm.ShowDialog()

#Load Locations Folders
if((Test-Path $customFoldersName) -eq $false)
{
    Set-DefaultLocations
}

New-ItemConditionalCreation -Item $AppPath -Type Directory
