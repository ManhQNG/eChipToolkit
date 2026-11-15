# =========================================================================
# TOOL FIX MACRO XLSM TO XLSX (DEEP CLEANING EXCEL)
# =========================================================================

$ErrorActionPreference = "Stop"

Write-Host "=====================================================" -ForegroundColor Green
Write-Host "          EXCEL MACRO CLEANER - PRO VERSION          " -ForegroundColor White
Write-Host "=====================================================" -ForegroundColor Green

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Khoi tao hop thoai chon thu muc (chi can khoi tao 1 lan)
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = "Chon thu muc chua cac file Excel .xlsm (Tool se quet ca thu muc con)"
$dialog.ShowNewFolderButton = $false
$dialog.RootFolder = [System.Environment+SpecialFolder]::MyComputer

# Ham ghi file XML an toan
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

# VONG LAP CHINH DE CHAY LIEN TUC
while ($true) {
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "`n[!] Da huy thao tac hoac thoat chuong trinh." -ForegroundColor Yellow
        break # Thoat khoi vong lap neu an Cancel
    }
    
    $root = $dialog.SelectedPath
    Write-Host "`n[+] Thu muc duoc chon: $root" -ForegroundColor Cyan

    $files = @(Get-ChildItem -LiteralPath $root -File -Filter "*.xlsm" -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Directory.Name -notmatch '(?i)^BACKUP$' })

    # XU LY KHI KHONG TIM THAY FILE
    if ($files.Count -eq 0) {
        Write-Host "`n[!] Khong tim thay file .xlsm nao hop le trong: $root" -ForegroundColor Yellow
        Write-Host "[?] Ban co muon chon phan vung/thu muc khac khong? (Y/N): " -NoNewline -ForegroundColor White
        $choice = Read-Host
        if ($choice -match '(?i)^n') {
            break # Thoat tool neu chon N
        } else {
            continue # Quay lai dau vong lap, mo lai hop thoai chon thu muc
        }
    }

    $success = 0
    $failed = 0

    foreach ($file in $files) {
        $source = $file.FullName
        $destination = [System.IO.Path]::ChangeExtension($source, ".xlsx")
        $zip = $null
        Write-Host "`nDang xu ly: $($file.Name)" -ForegroundColor Cyan

        if ([System.IO.File]::Exists($destination)) {
            Write-Host "  [-] BO QUA: Da ton tai file dich (Tranh ghi de)" -ForegroundColor DarkYellow
            $failed++
            continue
        }

        try {
            $backupDir = Join-Path $file.DirectoryName "BACKUP"
            if (-not ([System.IO.Directory]::Exists($backupDir))) {
                [System.IO.Directory]::CreateDirectory($backupDir) | Out-Null
            }
            
            $backupPath = Join-Path $backupDir ($file.BaseName + ".xlsm.BAK")
            if ([System.IO.File]::Exists($backupPath)) {
                $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $backupPath = Join-Path $backupDir ($file.BaseName + "_$stamp.xlsm.BAK")
            }

            [System.IO.File]::Copy($source, $backupPath, $true)
            
            $temp = Join-Path $file.DirectoryName ([Guid]::NewGuid().ToString("N") + ".tmp")
            [System.IO.File]::Copy($source, $temp, $true)

            try {
                $zip = [System.IO.Compression.ZipFile]::Open($temp, [System.IO.Compression.ZipArchiveMode]::Update)

                $removeNames = @($zip.Entries | Where-Object {
                    $_.FullName -match '(?i)(^|/)vbaProject\.bin$' -or
                    $_.FullName -match '(?i)(^|/)vbaData\.xml$'
                } | ForEach-Object { $_.FullName })
                foreach ($name in $removeNames) { $zip.GetEntry($name).Delete() }

                $ctEntry = $zip.GetEntry("[Content_Types].xml")
                if (-not $ctEntry) { throw "File khong dung chuan cau truc Excel." }
                
                $reader = New-Object System.IO.StreamReader($ctEntry.Open())
                [xml]$ctXml = $reader.ReadToEnd()
                $reader.Dispose()

                $nsMngr = New-Object System.Xml.XmlNamespaceManager($ctXml.NameTable)
                $nsMngr.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types")
                
                $vbaNodes = $ctXml.SelectNodes("//ct:Override[@PartName='/xl/vbaProject.bin' or @PartName='/xl/vbaData.xml']", $nsMngr)
                foreach ($node in $vbaNodes) { $node.ParentNode.RemoveChild($node) | Out-Null }

                $mainNode = $ctXml.SelectSingleNode("//ct:Override[@ContentType='application/vnd.ms-excel.sheet.macroEnabled.main+xml']", $nsMngr)
                if ($mainNode) { $mainNode.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") }
                
                Set-ZipEntryXml -Archive $zip -EntryName "[Content_Types].xml" -XmlData $ctXml

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

                [System.IO.File]::Move($temp, $destination)
                [System.IO.File]::Delete($source)
                
                Write-Host "  [+] THANH CONG: Da tao file an toan $($file.BaseName).xlsx" -ForegroundColor Green
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

    # TONG KET SAU KHI QUET XONG MOT THU MUC
    Write-Host "`n=====================================================" -ForegroundColor Green
    Write-Host " TONG KET XU LY" -ForegroundColor White
    Write-Host " Hoan tat thanh cong : $success file" -ForegroundColor Green
    Write-Host " Bi loi / Bo qua     : $failed file" -ForegroundColor Yellow
    Write-Host "=====================================================" -ForegroundColor Green
    
    Write-Host "`n[?] Ban co muon tiep tuc don dep phan vung khac? (Y/N): " -NoNewline -ForegroundColor Cyan
    $continueChoice = Read-Host
    if ($continueChoice -match '(?i)^n') {
        break # Thoat tool
    }
}