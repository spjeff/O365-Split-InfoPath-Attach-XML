<#
.SYNOPSIS
    Parse InfoPath attachment binary content saved as Base64 into new binary files.  Extract filename and file content for each InfoPath attachment XML node.  Save into subfolders and match original file naming.  Helpful for Office 365 migration and scenarios where InfoPath client is no longer available and users prefer to view attachments directly.

.EXAMPLE
	.\O365-Split-InfoPath-Attach-XML.ps1
    
.NOTES  
	File Name:  O365-Split-InfoPath-Attach-XML.ps1
	Author   :  Jeff Jones  - @spjeff

	Version  :  1.0.1
	Modified :  2020-04-28
#>

# Module
Import-Module "SharePointPnPPowerShellOnline" -ErrorAction SilentlyContinue | Out-Null

# Config
$webURL                     = "https://spjeff.sharepoint.com/sites/Team"
$sourceFormLibrary          = "/sites/team/AttachmentTest"
$destinationDocumentLibrary = "Shared Documents"

# Functions
function DownloadInfoPath() {
    # Enumerate source XMLs
    $folder = Get-PnPFolder -RelativeUrl $sourceFormLibrary
    Get-PnPProperty -ClientObject $folder -Property "Files" | Out-Null
    foreach ($file in $folder.Files) {
        # Download each XML
        Write-Host $file.Name
        Get-PnPFile -ServerRelativeUrl $file.ServerRelativeUrl -Path $destFolder -FileName $file.Name -AsFile
    }
}

function ParseInfoPath() {
    # Enumerate source attachments
    $files = Get-ChildItem $destFolder "*.xml"
    foreach ($f in $files) {
        # XPath parse attachment XML Base64 into Binary
        [xml]$xml = Get-Content $f.FullName
        ParseSingleXML $f.FullName $xml.myFields.field1
    }
}

function ParseSingleXML($xmlFileName, $attachBase64) {
    # Set text encoding
    $encoding = [System.Text.Encoding]::Unicode;
    $convert = [Convert]::FromBase64String($attachBase64)
    $ms = New-Object System.IO.MemoryStream(,$convert)
    $theReader = New-Object System.IO.BinaryReader($ms)
    
    # Parse file attachment name
    [System.Byte[]] $headerData = $theReader.ReadBytes(16);
    [Int]$fileSize = $theReader.ReadUInt32();
    [Int]$attachmentNameLength = $theReader.ReadUInt32() * 2;
    [System.Byte[]] $fileNameBytes = $theReader.ReadBytes($attachmentNameLength);
    [string]$fileName = $encoding.GetString($fileNameBytes, 0, $attachmentNameLength-2);
    
    # Write file content
    Write-Host "ATTACH $fileName" -Fore "Yellow"

    # Make folder
    $xmlFileName = $xmlFileName.Replace(".xml","_xml")
    mkdir $xmlFileName -ErrorAction SilentlyContinue | Out-Null
    
    # Write file
    $destFile = "$xmlFileName\$fileName"
    [IO.File]::WriteAllBytes($destFile, [Convert]::FromBase64String($attachBase64))
}
function UploadInfoPath() {
    # Enumerate source attachments
    $folders = Get-ChildItem $destFolder -Directory
    foreach ($folder in $folders) {
        # Create folder
        Write-Host "Create folder $($folder.Name)" -Fore "Green"
        $files = Get-ChildItem $folder.Fullname -File
        Add-PnPFolder -Name $folder.Name -Folder $destinationDocumentLibrary -ErrorAction SilentlyContinue | Out-Null

        # Loop files
        foreach ($file in $files) {
            # Upload each attachment
            Write-Host "Uploading $($file.Fullname)" -Fore "Green"
            $docLibPath = "$destinationDocumentLibrary/$($folder.Name)"
            $file = Add-PnPFile -Path  $file.Fullname -Folder $docLibPath
        }
    }
}

# Main
function Main() {
    # Connect
    $web = Get-PNPWeb
    if (!$web) {
        Connect-PnPOnline -Url $webURL -UseWebLogin
    }

    # Prepare folder
    $destFolder = $env:temp + "\O365-Split-InfoPath-Attach-XML"
    Remove-Item $destFolder -Confirm:$false -Force -Recurse
    New-Item $destFolder -Type "Directory" | Out-Null

    # Download
    DownloadInfoPath

    # Parse
    ParseInfoPath

    # Upload
    UploadInfoPath
}
Main