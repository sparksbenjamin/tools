<# 
.SYNOPSIS
    Convert all mail items in a PST file to individual PDF files with attachments.

.DESCRIPTION
    The script uses Outlook's COM interface to load a specified PST file, iterate through every folder,
    and save each email message as a PDF in a destination directory. Each email gets its own subfolder
    containing the PDF and all attachments.
    The script is fully self‑contained, performs robust error handling, and sanitises file names
    to avoid invalid characters. It also ensures that COM objects are released to prevent memory leaks.

.NOTES
    * Requires Microsoft Outlook to be installed on the machine where the script runs.
    * Outlook must be able to load the PST file without user interaction (i.e., the PST file
      is not password protected).
    * The script is written for PowerShell 5+ on Windows.

.PARAMETER PST
    Full path to the PST file to convert.

.PARAMETER Out
    Directory where PDF files will be written. The directory is created if it does not exist.

.EXAMPLE
    .\ConvertPSTtoPDF.ps1 -PST 'C:\Users\me\Documents\mail.pst' -Out 'C:\PDFs'

#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$PST,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$Out
)

# --------------------------------------------------------------------------- #
# Helper: Convert office documents to PDF using Word/Excel COM objects
# --------------------------------------------------------------------------- #
function ConvertTo-PDF {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLower()
    $converted = $false
    
    # Word documents
    if ($extension -in @('.doc', '.docx', '.rtf', '.txt', '.odt')) {
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Open($SourcePath)
            # wdFormatPDF = 17
            $doc.SaveAs([ref]$DestinationPath, [ref]17)
            $doc.Close()
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            $converted = $true
        } catch {
            Write-Warning "Failed to convert Word document '$SourcePath': $_"
        }
    }
    # Excel documents
    elseif ($extension -in @('.xls', '.xlsx', '.xlsm', '.csv')) {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($SourcePath)
            # xlTypePDF = 0
            $workbook.ExportAsFixedFormat(0, $DestinationPath)
            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            $converted = $true
        } catch {
            Write-Warning "Failed to convert Excel document '$SourcePath': $_"
        }
    }
    # PowerPoint documents
    elseif ($extension -in @('.ppt', '.pptx', '.pptm')) {
        try {
            $powerpoint = New-Object -ComObject PowerPoint.Application
            $presentation = $powerpoint.Presentations.Open($SourcePath, $true, $true, $false)
            # ppSaveAsPDF = 32
            $presentation.SaveAs($DestinationPath, 32)
            $presentation.Close()
            $powerpoint.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
            $converted = $true
        } catch {
            Write-Warning "Failed to convert PowerPoint document '$SourcePath': $_"
        }
    }
    
    return $converted
}

# --------------------------------------------------------------------------- #
# Helper: Sanitize file names so they can be used on Windows file system
# --------------------------------------------------------------------------- #
function ConvertTo-SafeFileName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [int]$MaxLength = 200
    )
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $safeName = $Name
    foreach ($char in $invalidChars) { 
        $safeName = $safeName -replace [Regex]::Escape($char), '_' 
    }
    # Trim trailing dots and spaces (Windows limitation)
    $safeName = $safeName.TrimEnd('. ')
    
    # Limit length to prevent path issues
    if ($safeName.Length -gt $MaxLength) {
        $safeName = $safeName.Substring(0, $MaxLength)
    }
    
    if ($safeName) { 
        return $safeName 
    } else { 
        return 'Untitled' 
    }
}

# --------------------------------------------------------------------------- #
# Helper: Generate unique folder name
# --------------------------------------------------------------------------- #
function Get-UniqueFolderPath {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )
    
    $safeFolderName = ConvertTo-SafeFileName -Name $FolderName
    $fullPath = Join-Path -Path $BasePath -ChildPath $safeFolderName
    $counter = 1
    
    while (Test-Path $fullPath) {
        $uniqueName = "{0}_{1}" -f $safeFolderName, $counter
        $fullPath = Join-Path -Path $BasePath -ChildPath $uniqueName
        $counter++
    }
    
    return $fullPath
}

# --------------------------------------------------------------------------- #
# Main conversion logic
# --------------------------------------------------------------------------- #
function Convert-PSTToPDF {
    param (
        [string]$PSTPath,
        [string]$OutputDirectory
    )

    # Ensure output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        try {
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
            Write-Host "Created output directory: $OutputDirectory" -ForegroundColor Cyan
        } catch {
            throw "Failed to create output directory: $_"
        }
    }

    # Create Outlook COM objects
    $outlook = $null
    $namespace = $null
    $store = $null
    
    try {
        Write-Host "Starting Outlook..." -ForegroundColor Cyan
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace('MAPI')
        $namespace.Logon('', '', $false, $false) | Out-Null
        Write-Host "Outlook initialized successfully." -ForegroundColor Green
    } catch {
        throw "Outlook could not be started or logged on: $_"
    }

    try {
        # Add the PST file if it isn't already loaded
        Write-Host "Loading PST file: $PSTPath" -ForegroundColor Cyan
        $existingStore = $namespace.Stores | Where-Object { $_.FilePath -eq $PSTPath }
        
        if ($null -eq $existingStore) {
            $namespace.AddStore($PSTPath)
            Start-Sleep -Seconds 2  # Allow time for the store to be added
            $store = $namespace.Stores | Where-Object { $_.FilePath -eq $PSTPath }
        } else {
            $store = $existingStore
        }

        if ($null -eq $store) {
            throw "Could not load PST file. Please verify the file path and that Outlook can access it."
        }

        Write-Host "PST file loaded successfully." -ForegroundColor Green
        
        # Initialize counters
        $script:totalEmails = 0
        $script:successCount = 0
        $script:failCount = 0
        $script:attachmentCount = 0
        $script:convertedAttachments = 0

        # Process all folders recursively
        $rootFolder = $store.GetRootFolder()
        Write-Host "`nProcessing emails..." -ForegroundColor Cyan
        Process-Folder -Folder $rootFolder -OutputDir $OutputDirectory
        
        Write-Host "`n========================================" -ForegroundColor Green
        Write-Host "Conversion Complete!" -ForegroundColor Green
        Write-Host "Total emails processed: $script:totalEmails" -ForegroundColor White
        Write-Host "Successfully converted: $script:successCount" -ForegroundColor Green
        Write-Host "Failed conversions: $script:failCount" -ForegroundColor $(if($script:failCount -gt 0){'Yellow'}else{'Green'})
        Write-Host "Total attachments saved: $script:attachmentCount" -ForegroundColor White
        Write-Host "Attachments converted to PDF: $script:convertedAttachments" -ForegroundColor White
        Write-Host "========================================" -ForegroundColor Green
        
    } catch {
        throw "An error occurred during conversion: $_"
    } finally {
        # Clean up: remove the PST store if it was added by this script
        if ($null -ne $store -and $null -eq $existingStore) {
            try { 
                Write-Host "`nRemoving PST store..." -ForegroundColor Cyan
                $namespace.RemoveStore($store.GetRootFolder())
            } catch { 
                Write-Warning "Could not remove PST store: $_"
            }
        }

        # Release COM objects
        if ($null -ne $namespace) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
        if ($null -ne $outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Cleanup complete." -ForegroundColor Cyan
    }
}

function Process-Folder {
    param (
        [Parameter(Mandatory = $true)]
        $Folder,

        [Parameter(Mandatory = $true)]
        [string]$OutputDir
    )

    try {
        $folderName = $Folder.Name
        Write-Host "Processing folder: $folderName" -ForegroundColor Yellow
        
        # Process each mail item in the current folder
        $itemCount = $Folder.Items.Count
        $currentItem = 0
        
        foreach ($item in $Folder.Items) {
            $currentItem++
            
            # Check if item is a MailItem (type 43 = olMailItem)
            if ($item.Class -eq 43) {
                $script:totalEmails++
                Write-Progress -Activity "Processing Emails" -Status "Folder: $folderName" -CurrentOperation "Email $currentItem of $itemCount" -PercentComplete (($currentItem / $itemCount) * 100)
                Save-MailAsPDF -MailItem $item -OutputDir $OutputDir
            }
            
            # Release the COM object for this item
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null
        }
        
        Write-Progress -Activity "Processing Emails" -Completed

        # Recursively process sub‑folders
        foreach ($subFolder in $Folder.Folders) {
            Process-Folder -Folder $subFolder -OutputDir $OutputDir
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($subFolder) | Out-Null
        }
        
    } catch {
        Write-Warning "Error processing folder '$($Folder.Name)': $_"
    }
}

function Save-MailAsPDF {
    param (
        [Parameter(Mandatory = $true)]
        $MailItem,

        [Parameter(Mandatory = $true)]
        [string]$OutputDir
    )

    $subject = if ($MailItem.Subject) { $MailItem.Subject } else { "No Subject" }
    $receivedTime = $MailItem.ReceivedTime
    $timestamp = $receivedTime.ToString("yyyy-MM-dd_HHmmss")
    
    # Create unique folder name with timestamp and subject
    $folderName = "{0}_{1}" -f $timestamp, $subject
    $emailFolder = Get-UniqueFolderPath -BasePath $OutputDir -FolderName $folderName
    
    try {
        # Create folder for this email
        New-Item -ItemType Directory -Path $emailFolder -Force | Out-Null
        
        # Save email as PDF
        $pdfFileName = "Email.pdf"
        $pdfPath = Join-Path -Path $emailFolder -ChildPath $pdfFileName
        
        # olFormatPDF = 17
        $MailItem.SaveAs($pdfPath, 17)
        
        # Save attachments
        $attachmentsSaved = 0
        try {
            $attachmentCount = $MailItem.Attachments.Count
            if ($attachmentCount -gt 0) {
                for ($i = 1; $i -le $attachmentCount; $i++) {
                    try {
                        $attachment = $MailItem.Attachments.Item($i)
                        $attachmentName = ConvertTo-SafeFileName -Name $attachment.FileName
                        $attachmentPath = Join-Path -Path $emailFolder -ChildPath $attachmentName
                        
                        # Handle duplicate attachment names
                        $counter = 1
                        while (Test-Path $attachmentPath) {
                            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($attachmentName)
                            $extension = [System.IO.Path]::GetExtension($attachmentName)
                            $attachmentName = "{0}_{1}{2}" -f $baseName, $counter, $extension
                            $attachmentPath = Join-Path -Path $emailFolder -ChildPath $attachmentName
                            $counter++
                        }
                        
                        $attachment.SaveAsFile($attachmentPath)
                        $attachmentsSaved++
                        $script:attachmentCount++
                        
                        # Try to convert to PDF if it's an Office document
                        $pdfAttachmentPath = [System.IO.Path]::ChangeExtension($attachmentPath, '.pdf')
                        if (ConvertTo-PDF -SourcePath $attachmentPath -DestinationPath $pdfAttachmentPath) {
                            # Delete original file after successful conversion
                            Remove-Item -Path $attachmentPath -Force
                            $script:convertedAttachments++
                        }
                        
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($attachment) | Out-Null
                    } catch {
                        Write-Warning "Failed to save attachment from email '$subject': $_"
                    }
                }
            }
        } catch {
            Write-Warning "Error processing attachments for email '$subject': $_"
        }
        
        $script:successCount++
        $attachmentInfo = if ($attachmentsSaved -gt 0) { " [$attachmentsSaved attachment(s)]" } else { "" }
        Write-Host "  ✓ Saved: $subject$attachmentInfo" -ForegroundColor Gray
        
    } catch {
        $script:failCount++
        Write-Warning "Failed to save email '$subject': $_"
    }
}

# --------------------------------------------------------------------------- #
# Execute script
# --------------------------------------------------------------------------- #
if ($PSCmdlet.ShouldProcess("$PST", 'Convert to PDF')) {
    Convert-PSTToPDF -PSTPath $PST -OutputDirectory $Out
}
