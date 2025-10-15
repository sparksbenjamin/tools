<# 
.SYNOPSIS
    Convert all mail items in a PST file to individual PDF files with attachments.

.DESCRIPTION
    Loads a PST file into Outlook, walks all folders, and saves each email to PDF with
    attachments in its own subfolder. Office documents (Word, Excel, PowerPoint) are
    automatically converted to PDF.

.NOTES
    * Requires Microsoft Outlook + Word/Excel/PowerPoint installed.
    * PST must not be password protected.
    * Written for PowerShell 5+ on Windows.

.EXAMPLE
    .\ConvertPSTtoPDF.ps1 -PST 'C:\Users\me\Documents\mail.pst' -Out 'C:\PDFs'
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$PST,

    [Parameter(Mandatory = $true)]
    [string]$Out
)

# --------------------------------------------------------------------------- #
# Helper: Convert office documents to PDF using Word/Excel/PowerPoint COM
# --------------------------------------------------------------------------- #
function ConvertTo-PDF {
    param (
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )
    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLower()
    $converted = $false
    $retries = 3

    for ($try = 1; $try -le $retries -and -not $converted; $try++) {
        try {
            switch ($extension) {
                {$_ -in '.doc', '.docx', '.rtf', '.txt', '.odt'} {
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $doc = $word.Documents.Open($SourcePath)
                    $doc.SaveAs([ref]$DestinationPath, [ref]17)  # wdFormatPDF
                    $doc.Close()
                    $converted = $true
                }
                {$_ -in '.xls', '.xlsx', '.xlsm', '.csv'} {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                    $workbook = $excel.Workbooks.Open($SourcePath)
                    $workbook.ExportAsFixedFormat(0, $DestinationPath) # xlTypePDF
                    $workbook.Close($false)
                    $converted = $true
                }
                {$_ -in '.ppt', '.pptx', '.pptm'} {
                    $powerpoint = New-Object -ComObject PowerPoint.Application
                    $powerpoint.Visible = $false
                    $powerpoint.DisplayAlerts = 0
                    $presentation = $powerpoint.Presentations.Open($SourcePath, $true, $true, $false)
                    $presentation.SaveAs($DestinationPath, 32) # ppSaveAsPDF
                    $presentation.Close()
                    $converted = $true
                }
            }
        } catch {
            if ($try -lt $retries) { Start-Sleep -Milliseconds 300 } 
            else { Write-Warning "Failed to convert '$SourcePath': $_" }
        } finally {
            if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc); $doc = $null }
            if ($word) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word); $word = $null }
            if ($workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook); $workbook = $null }
            if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel); $excel = $null }
            if ($presentation) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation); $presentation = $null }
            if ($powerpoint) { $powerpoint.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint); $powerpoint = $null }
        }
    }
    return $converted
}

# --------------------------------------------------------------------------- #
# Helper: Sanitize file names for Windows
# --------------------------------------------------------------------------- #
function ConvertTo-SafeFileName {
    param ([Parameter(Mandatory = $true)][string]$Name, [int]$MaxLength = 120)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = $Name
    foreach ($c in $invalid) { $safe = $safe -replace [Regex]::Escape($c), '_' }
    $safe = $safe.TrimEnd('. ')
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength) }
    return ($safe -ne '') ? $safe : 'Untitled'
}

# --------------------------------------------------------------------------- #
# Helper: Generate unique folder name
# --------------------------------------------------------------------------- #
function Get-UniqueFolderPath {
    param ([string]$BasePath, [string]$FolderName)
    $safe = ConvertTo-SafeFileName -Name $FolderName
    $path = Join-Path $BasePath $safe
    $i = 1
    while (Test-Path $path) {
        $path = Join-Path $BasePath ("{0}_{1}" -f $safe, $i)
        $i++
    }
    return $path
}

# --------------------------------------------------------------------------- #
# Main: Convert PST
# --------------------------------------------------------------------------- #
function Convert-PSTToPDF {
    param ([string]$PSTPath, [string]$OutputDirectory)

    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }

    $outlook = $null; $namespace = $null; $store = $null
    try {
        Write-Host "Starting Outlook..." -ForegroundColor Cyan
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace('MAPI')
        $namespace.Logon('', '', $false, $false) | Out-Null
    } catch { throw "Failed to start Outlook: $_" }

    try {
        Write-Host "Loading PST file: $PSTPath" -ForegroundColor Cyan
        $existingStore = $namespace.Stores | Where-Object { $_.FilePath -eq $PSTPath }
        if (-not $existingStore) {
            $namespace.AddStore($PSTPath)
            Start-Sleep 2
            $store = $namespace.Stores | Where-Object { $_.FilePath -eq $PSTPath }
        } else { $store = $existingStore }

        if (-not $store) { throw "Could not load PST file." }

        Write-Host "PST file loaded successfully." -ForegroundColor Green
        $script:totalEmails = 0
        $script:successCount = 0
        $script:failCount = 0
        $script:attachmentCount = 0
        $script:convertedAttachments = 0

        $root = $store.GetRootFolder()
        Write-Host "`nProcessing emails..." -ForegroundColor Cyan
        Process-Folder -Folder $root -OutputDir $OutputDirectory

        Write-Host "`n========= SUMMARY =========" -ForegroundColor Green
        Write-Host "Emails processed: $script:totalEmails"
        Write-Host "Converted successfully: $script:successCount"
        Write-Host "Failed conversions: $script:failCount"
        Write-Host "Attachments saved: $script:attachmentCount"
        Write-Host "Attachments converted: $script:convertedAttachments"
        Write-Host "===========================" -ForegroundColor Green
    } finally {
        if ($store -and -not $existingStore) {
            try {
                Write-Host "Removing PST store..." -ForegroundColor Cyan
                $namespace.RemoveStore($store.GetRootFolder())
            } catch { Write-Warning "Failed to remove PST store: $_" }
        }
        if ($namespace) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) }
        if ($outlook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        Write-Host "Cleanup complete." -ForegroundColor Cyan
    }
}

# --------------------------------------------------------------------------- #
# Folder Processing
# --------------------------------------------------------------------------- #
function Process-Folder {
    param ($Folder, [string]$OutputDir)
    try {
        $folderName = $Folder.Name
        Write-Host "Processing folder: $folderName" -ForegroundColor Yellow

        $items = @($Folder.Items)
        $itemCount = $items.Count
        $current = 0

        foreach ($item in $items) {
            $current++
            if ($item.Class -eq 43) {  # olMailItem
                $script:totalEmails++
                Write-Progress -Activity "Processing Emails" `
                    -Status "Folder: $folderName" `
                    -CurrentOperation "Email $current of $itemCount" `
                    -PercentComplete (($current / $itemCount) * 100)
                Save-MailAsPDF -MailItem $item -OutputDir $OutputDir
            }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item)
        }

        Write-Progress -Activity "Processing Emails" -Completed

        foreach ($sub in $Folder.Folders) {
            Process-Folder -Folder $sub -OutputDir $OutputDir
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sub)
        }
    } catch {
        Write-Warning "Error processing folder '$($Folder.Name)': $_"
    } finally {
        if ($items) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) }
    }
}

# --------------------------------------------------------------------------- #
# Save Mail as PDF and export attachments
# --------------------------------------------------------------------------- #
function Save-MailAsPDF {
    param ($MailItem, [string]$OutputDir)

    $subject = if ($MailItem.Subject) { $MailItem.Subject } else { "No Subject" }
    $timestamp = $MailItem.ReceivedTime.ToString("yyyy-MM-dd_HHmmss")
    $folderName = "{0}_{1}" -f $timestamp, $subject
    $emailFolder = Get-UniqueFolderPath -BasePath $OutputDir -FolderName $folderName

    try {
        New-Item -ItemType Directory -Path $emailFolder -Force | Out-Null
        $pdfPath = Join-Path $emailFolder "Email.pdf"

        # Save MailItem as MSG then convert via Word to PDF
        $tempMsg = Join-Path $emailFolder "temp.msg"
        $MailItem.SaveAs($tempMsg, 3)  # olMSGUnicode

        $word = $null; $doc = $null
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Open($tempMsg)
            $doc.SaveAs([ref]$pdfPath, [ref]17) # wdFormatPDF
            $doc.Close()
        } finally {
            if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
            if ($word) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
            Remove-Item $tempMsg -Force
        }

        # Save attachments
        $attachmentsSaved = 0
        $attachmentCount = $MailItem.Attachments.Count
        if ($attachmentCount -gt 0) {
            for ($i = 1; $i -le $attachmentCount; $i++) {
                try {
                    $att = $MailItem.Attachments.Item($i)
                    $attName = ConvertTo-SafeFileName -Name $att.FileName
                    $attPath = Join-Path $emailFolder $attName

                    $counter = 1
                    while (Test-Path $attPath) {
                        $base = [System.IO.Path]::GetFileNameWithoutExtension($attName)
                        $ext = [System.IO.Path]::GetExtension($attName)
                        $attName = "{0}_{1}{2}" -f $base, $counter, $ext
                        $attPath = Join-Path $emailFolder $attName
                        $counter++
                    }

                    $att.SaveAsFile($attPath)
                    $attachmentsSaved++
                    $script:attachmentCount++

                    $ext = [System.IO.Path]::GetExtension($attPath).ToLower()
                    if ($ext -ne '.pdf') {
                        $pdfAttPath = [System.IO.Path]::ChangeExtension($attPath, '.pdf')
                        if (ConvertTo-PDF -SourcePath $attPath -DestinationPath $pdfAttPath) {
                            Remove-Item $attPath -Force
                            $script:convertedAttachments++
                        }
                    }

                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($att)
                } catch {
                    Write-Warning "Failed to save attachment for '$subject': $_"
                }
            }
        }

        $script:successCount++
        $msg = if ($attachmentsSaved -gt 0) { " [$attachmentsSaved attachment(s)]" } else { "" }
        Write-Host "  ✓ Saved: $subject$msg" -ForegroundColor Gray

    } catch {
        $script:failCount++
        Write-Warning "Failed to save email '$subject': $_"
    }
}

# --------------------------------------------------------------------------- #
# Execute
# --------------------------------------------------------------------------- #
if ($PSCmdlet.ShouldProcess($PST, 'Convert to PDF')) {
    Convert-PSTToPDF -PSTPath $PST -OutputDirectory $Out
}
