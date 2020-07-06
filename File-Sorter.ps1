<#
.SYNOPSIS
  Moves files from a source folder to Videos, Pictures, Documents and Program Installers folders based on file extensions.
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
  Creation Date:  30.06.2020.
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Set known extensions
$documentExtensions = "*.DOC", "*.DOCX", "*.HTML", "*.HTM", "*.ODT", "*.PDF", "*.XLS", "*.XLSX", "*.ODS", "*.PPT", "*.PPTX", "*.TXT", "*.LOG"
$installerExtensions = "*.exe", "*.msi", "*.msm", "*.msp", "*.mst", "*.idt", "*.idt", "*.cub", "*.pcp", "*.jar"
$pictureExtensions = "*.JPG", "*.PNG", "*.GIF", "*.WEBP", "*.TIFF", "*.SD", "*.RAW", "*.BMP", "*.HEIF", "*.INDD", "*.JPEG", "*.SVG", "*.AI", "*.EPS", "*.PDF", "*.cvs"
$videoExtensions = "*.WEBM", "*.MPG", "*.MP2", "*.MPEG", "*.MPE", "*.MPV", "*.OGG", "*.MP4", "*.M4P", "*.M4V", "*.AVI", "*.WMV", "*.MO", "*.QT", "*.FLV", "*.SWF", "*.AVCHD"

#Create GUI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$formFileSorter                  = New-Object system.Windows.Forms.Form
$formFileSorter.ClientSize       = New-Object System.Drawing.Point(795,325)
$formFileSorter.text             = "File Sorter"
$formFileSorter.TopMost          = $true

$lblDownloads                    = New-Object system.Windows.Forms.Label
$lblDownloads.text               = "Downloads"
$lblDownloads.AutoSize           = $true
$lblDownloads.width              = 25
$lblDownloads.height             = 10
$lblDownloads.location           = New-Object System.Drawing.Point(20,64)
$lblDownloads.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtDownloads                    = New-Object system.Windows.Forms.TextBox
$txtDownloads.multiline          = $false
$txtDownloads.width              = 600
$txtDownloads.height             = 20
$txtDownloads.location           = New-Object System.Drawing.Point(170,60)
$txtDownloads.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblDocuments                    = New-Object system.Windows.Forms.Label
$lblDocuments.text               = "Documents"
$lblDocuments.AutoSize           = $true
$lblDocuments.width              = 25
$lblDocuments.height             = 10
$lblDocuments.location           = New-Object System.Drawing.Point(20,137)
$lblDocuments.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtDocuments                    = New-Object system.Windows.Forms.TextBox
$txtDocuments.multiline          = $false
$txtDocuments.width              = 600
$txtDocuments.height             = 20
$txtDocuments.location           = New-Object System.Drawing.Point(170,133)
$txtDocuments.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblPictures                     = New-Object system.Windows.Forms.Label
$lblPictures.text                = "Pictures"
$lblPictures.AutoSize            = $true
$lblPictures.width               = 25
$lblPictures.height              = 10
$lblPictures.location            = New-Object System.Drawing.Point(20,167)
$lblPictures.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtPictures                     = New-Object system.Windows.Forms.TextBox
$txtPictures.multiline           = $false
$txtPictures.width               = 600
$txtPictures.height              = 20
$txtPictures.location            = New-Object System.Drawing.Point(170,163)
$txtPictures.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblVideos                       = New-Object system.Windows.Forms.Label
$lblVideos.text                  = "Videos"
$lblVideos.AutoSize              = $true
$lblVideos.width                 = 25
$lblVideos.height                = 10
$lblVideos.location              = New-Object System.Drawing.Point(20,197)
$lblVideos.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtVideos                       = New-Object system.Windows.Forms.TextBox
$txtVideos.multiline             = $false
$txtVideos.width                 = 600
$txtVideos.height                = 20
$txtVideos.location              = New-Object System.Drawing.Point(170,194)
$txtVideos.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnSortFiles                    = New-Object system.Windows.Forms.Button
$btnSortFiles.text               = "Sort Files"
$btnSortFiles.width              = 150
$btnSortFiles.height             = 30
$btnSortFiles.location           = New-Object System.Drawing.Point(327,268)
$btnSortFiles.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10,
                                    [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$lblProgramInstallers            = New-Object system.Windows.Forms.Label
$lblProgramInstallers.text       = "Program Installers"
$lblProgramInstallers.AutoSize   = $true
$lblProgramInstallers.width      = 25
$lblProgramInstallers.height     = 10
$lblProgramInstallers.location   = New-Object System.Drawing.Point(21,229)
$lblProgramInstallers.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtProgramInstallers            = New-Object system.Windows.Forms.TextBox
$txtProgramInstallers.multiline  = $false
$txtProgramInstallers.width      = 600
$txtProgramInstallers.height     = 20
$txtProgramInstallers.location   = New-Object System.Drawing.Point(170,223)
$txtProgramInstallers.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnDefaultLocations             = New-Object system.Windows.Forms.Button
$btnDefaultLocations.text        = "Default Locations"
$btnDefaultLocations.width       = 150
$btnDefaultLocations.height      = 30
$btnDefaultLocations.location    = New-Object System.Drawing.Point(52,268)
$btnDefaultLocations.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnSaveLocations                = New-Object system.Windows.Forms.Button
$btnSaveLocations.text           = "Save Locations"
$btnSaveLocations.width          = 150
$btnSaveLocations.height         = 30
$btnSaveLocations.location       = New-Object System.Drawing.Point(592,268)
$btnSaveLocations.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblSourceFolder                 = New-Object system.Windows.Forms.Label
$lblSourceFolder.text            = "Source Folder"
$lblSourceFolder.AutoSize        = $true
$lblSourceFolder.width           = 50
$lblSourceFolder.height          = 10
$lblSourceFolder.location        = New-Object System.Drawing.Point(405,24)
$lblSourceFolder.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,
                                    [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$lblDestinationFolders           = New-Object system.Windows.Forms.Label
$lblDestinationFolders.text      = "Destination Folders"
$lblDestinationFolders.AutoSize  = $true
$lblDestinationFolders.width     = 50
$lblDestinationFolders.height    = 10
$lblDestinationFolders.location  = New-Object System.Drawing.Point(388,102)
$lblDestinationFolders.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif',10,
                                    [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$formFileSorter.controls.AddRange(@($lblDownloads,$txtDownloads,$lblDocuments,$txtDocuments,$lblPictures,$txtPictures,$lblVideos,
$txtVideos,$btnSortFiles,$lblProgramInstallers,$txtProgramInstallers,$btnDefaultLocations,$btnSaveLocations,$lblSourceFolder,
$lblDestinationFolders))

#Adding events to buttons
$btnSortFiles.Add_Click({ Start-FileSorting })
$btnDefaultLocations.Add_Click({ Set-DefaultLocations })
$btnSaveLocations.Add_Click({ Save-FolderSettings })

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Default Paths
$defaultSourceFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$defaultPicturesFolder = [environment]::getfolderpath("mypictures")
$defaultProgramInstallersFolder = "D:\Program Installers\"
$defaultDocumentsFolder = [environment]::getfolderpath("mydocuments")
$defaultVideosFolder = [environment]::getfolderpath("myvideos")

#Aplication Folder Info
$appPath = $env:APPDATA + "\File Sorter"

#Log File Info
$logName = "File-Sorter-Log.txt"
$logFile = Join-Path -Path $appPath -ChildPath $logName

#Custom Folder Location File Info
$customFoldersName = "Custom-Folders.xml"
$customFoldersFile = Join-Path -Path $appPath -ChildPath $customFoldersName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

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
New-ItemConditionalCreation -Item "D:\Test.txt" -Type File

.NOTES

#>
function New-ItemConditionalCreation
{
    param([String]$Item, [String]$Type)

    if((Test-Path $Item) -eq $false)
    {
        New-Item -Path $Item -ItemType $Type
    }
}

<#
.SYNOPSIS
Writes a log entry

.DESCRIPTION
Creates a log entry with timestamp and message passed thru a parameter $Massage, and saves the log entry to log file in
%APPDATA%\File Sorter\File Sorter Log.txt

.PARAMETER Message
String value to be writen in the log file alongside timestamp

.EXAMPLE
Write-Log -Message "File sorting started"

.NOTES
Format of the timestamp in "yyyy.MM.dd. HH:mm:ss" and this function adds " - " after timestamp and before the main massage.
#>
function Write-Log
{
    param([String]$Message)

    $timestamp = Get-Date -Format "yyyy.MM.dd. HH:mm:ss"
    $logEntry = $timestamp + " - " + $Message
    Add-content $logfile -Value $logEntry
}

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
    if((Test-Path -Path $customFoldersFile) -eq $false)
    {
        #Set The Formatting
        $xmlsettings = New-Object System.Xml.XmlWriterSettings
        $xmlsettings.Indent = $true
        $xmlsettings.IndentChars = "    "

        #Set the File Name Create The Document
        $XmlWriter = [System.XML.XmlWriter]::Create($customFoldersFile, $xmlsettings)

        #Write the XML Decleration and set the XSL
        $xmlWriter.WriteStartDocument()
        $xmlWriter.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='style.xsl'")

        #Start the Root Element
        $xmlWriter.WriteStartElement("Root")

        Save-Path -Name "Source" -Path $txtDownloads.Text
        Save-Path -Name "Documents" -Path $txtDocuments.Text
        Save-Path -Name "Pictures" -Path $txtPictures.Text
        Save-Path -Name "Videos" -Path $txtVideos.Text
        Save-Path -Name "ProgramInstallers" -Path $txtProgramInstallers.Text

        $xmlWriter.WriteEndElement()

        #End, Finalize and close the XML Document
        $xmlWriter.WriteEndDocument()
        $xmlWriter.Flush()
        $xmlWriter.Close()
    }

    #TODO Update saved paths
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

#Load Locations Folders
if((Test-Path $customFoldersName) -eq $false)
{
    Set-DefaultLocations
}

New-ItemConditionalCreation -Item $appPath -Type Directory
New-ItemConditionalCreation -Item $logFile -Type File

#Launch Window
[void]$formFileSorter.ShowDialog()