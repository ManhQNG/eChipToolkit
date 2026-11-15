# =========================================================================
# TOOL FIX MACRO DOCM TO DOCX (DEEP CLEANING)
# =========================================================================

$ErrorActionPreference = "Stop"

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "          WORD MACRO CLEANER - PRO VERSION           " -ForegroundColor White
Write-Host "=====================================================" -ForegroundColor Cyan

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Hop thoai chon thu muc
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = "Chon thu muc chua cac file Word .docm (Tool se quet ca thu muc con)"
$dialog.ShowNewFolderButton = $false
$dialog.RootFolder = [System.Environment+SpecialFolder]::MyComputer

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "`n[!] Da huy thao tac." -ForegroundColor Yellow
    exit 0
}
$root = $dialog.SelectedPath
Write-Host "`n[+] Thu muc goc: $root" -ForegroundColor Green

# Loc file DOCM, loai tru chinh xac cac thu muc co ten la "BACKUP" (tranh chan nham path)
$files = @(Get-ChildItem -LiteralPath $root -File -Filter "*.docm" -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.Directory.Name -notmatch '(?i)^BACKUP$' })

if ($files.Count -eq 0) {
    Write-Host "`n[!] Khong tim thay file .docm nao hop le de xu ly." -ForegroundColor Yellow
    Read-Host "Nhan Enter de thoat"
    exit 0
}

# Ham ghi file XML an toan khong co BOM (Byte Order Mark) de tranh loi Word
function Set-ZipEntryXml {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [xml]$XmlData
    )
    $old = $Archive.GetEntry($EntryName)
    if ($old) { $old.Delete() }
    
    $new = $Archive.CreateEntry($EntryName)
    $stream = $new.Open()
    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.Xml.XmlTextWriter($stream, $utf8NoBom)
        $writer.Formatting = [System.Xml.Formatting]::None
        $XmlData.Save($writer)
        $writer.Flush()
    } finally {
        if ($writer) { $writer.Dispose() }
        $stream.Dispose()
    }
}

$success = 0
$failed = 0

foreach ($file in $files) {
    $source = $file.FullName
    $destination = [System.IO.Path]::ChangeExtension($source, ".docx")
    $zip = $null
    Write-Host "`nDang xu ly: $($file.Name)" -ForegroundColor Cyan

    if ([System.IO.File]::Exists($destination)) {
        Write-Host "  [-] BO QUA: Da ton tai file dich (Tranh ghi de)" -ForegroundColor DarkYellow
        $failed++
        continue
    }

    try {
        # Tao thu muc BACKUP
        $backupDir = Join-Path $file.DirectoryName "BACKUP"
        if (-not ([System.IO.Directory]::Exists($backupDir))) {
            [System.IO.Directory]::CreateDirectory($backupDir) | Out-Null
        }
        
        $backupPath = Join-Path $backupDir ($file.BaseName + ".docm.BAK")
        if ([System.IO.File]::Exists($backupPath)) {
            $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $backupPath = Join-Path $backupDir ($file.BaseName + "_$stamp.docm.BAK")
        }

        # Su dung API .NET de vuot rao gioi han 260 ky tu duong dan cua Windows
        [System.IO.File]::Copy($source, $backupPath, $true)
        
        $temp = Join-Path $file.DirectoryName ([Guid]::NewGuid().ToString("N") + ".tmp")
        [System.IO.File]::Copy($source, $temp, $true)

        try {
            $zip = [System.IO.Compression.ZipFile]::Open($temp, [System.IO.Compression.ZipArchiveMode]::Update)

            # 1. Xoa file bin/xml chua ma doc/VBA
            $removeNames = @($zip.Entries | Where-Object {
                $_.FullName -match '(?i)(^|/)vbaProject\.bin$' -or
                $_.FullName -match '(?i)(^|/)vbaData\.xml$'
            } | ForEach-Object { $_.FullName })
            foreach ($name in $removeNames) { $zip.GetEntry($name).Delete() }

            # 2. Xu ly [Content_Types].xml bang chuan XML DOM
            $ctEntry = $zip.GetEntry("[Content_Types].xml")
            if (-not $ctEntry) { throw "File khong dung chuan cau truc Word." }
            
            $reader = New-Object System.IO.StreamReader($ctEntry.Open())
            [xml]$ctXml = $reader.ReadToEnd()
            $reader.Dispose()

            $nsMngr = New-Object System.Xml.XmlNamespaceManager($ctXml.NameTable)
            $nsMngr.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types")
            
            # Xoa khai bao vbaProject
            $vbaNodes = $ctXml.SelectNodes("//ct:Override[@PartName='/word/vbaProject.bin' or @PartName='/word/vbaData.xml']", $nsMngr)
            foreach ($node in $vbaNodes) { $node.ParentNode.RemoveChild($node) | Out-Null }

            # Doi Type tai lieu sang chuan DOCX
            $mainNode = $ctXml.SelectSingleNode("//ct:Override[@ContentType='application/vnd.ms-word.document.macroEnabled.main+xml']", $nsMngr)
            if ($mainNode) { $mainNode.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml") }
            
            Set-ZipEntryXml -Archive $zip -EntryName "[Content_Types].xml" -XmlData $ctXml

            # 3. Quet va don dep moi file .rels bang chuan XML DOM
            $relsNames = @($zip.Entries | Where-Object { $_.FullName -match '(?i)\.rels$' } | ForEach-Object { $_.FullName })
            foreach ($relName in $relsNames) {
                $relEntry = $zip.GetEntry($relName)
                $reader = New-Object System.IO.StreamReader($relEntry.Open())
                [xml]$relXml = $reader.ReadToEnd()
                $reader.Dispose()

                $relNs = New-Object System.Xml.XmlNamespaceManager($relXml.NameTable)
                $relNs.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")

                $badRels = $relXml.SelectNodes("//r:Relationship[@Target='vbaProject.bin' or @Target='vbaData.xml']", $relNs)
                if ($badRels.Count -gt 0) {
                    foreach ($node in $badRels) { $node.ParentNode.RemoveChild($node) | Out-Null }
                    Set-ZipEntryXml -Archive $zip -EntryName $relName -XmlData $relXml
                }
            }

            $zip.Dispose(); $zip = $null

            # Hoan tat: Doi ten thanh file dich va xoa file goc
            [System.IO.File]::Move($temp, $destination)
            [System.IO.File]::Delete($source)
            
            Write-Host "  [+] THANH CONG: Da tao file an toan $($file.BaseName).docx" -ForegroundColor Green
            $success++
        }
        finally {
            if ($zip) { $zip.Dispose() }
            if ([System.IO.File]::Exists($temp)) { [System.IO.File]::Delete($temp) }
        }
    }
    catch {
        Write-Host "  [!] LOI: $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host " TONG KET XU LY" -ForegroundColor White
Write-Host " Hoan tat thanh cong : $success file" -ForegroundColor Green
Write-Host " Bi loi / Bo qua     : $failed file" -ForegroundColor Yellow
Write-Host "=====================================================" -ForegroundColor Cyan
Read-Host "Nhan Enter de ket thuc"