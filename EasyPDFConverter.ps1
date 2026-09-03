<#
.SYNOPSIS
    EasyPDFConverter - конвертирует PDF-файлы в PNG или JPEG в максимальном качестве.

.DESCRIPTION
    Скрипт берёт все PDF-файлы из папки, в которой лежит, спрашивает формат и разрешение,
    и сохраняет каждую страницу отдельной картинкой в папку output\<имя PDF>\.

    Рендеринг выполняет pdftoppm из проекта Poppler. При первом запуске портативная сборка
    Poppler для Windows автоматически скачивается в папку tools\poppler - ничего
    устанавливать не нужно.

    PNG сохраняется без потерь. JPEG - с качеством 100 и оптимизацией таблиц Хаффмана.

.PARAMETER Path
    Один или несколько PDF-файлов либо папок с ними. Если не указано - берутся все PDF
    из папки скрипта. Сюда же попадают файлы, перетащенные мышкой на .bat-лаунчер.

.PARAMETER Format
    png или jpeg. Если не указан - скрипт спросит.

.PARAMETER Dpi
    Разрешение рендеринга в точках на дюйм (36-2400). Если не указано - скрипт спросит.

.PARAMETER NoPause
    Не ждать нажатия клавиши в конце. Используется .bat-лаунчером.

.PARAMETER NoOpen
    Не открывать папку output в Проводнике после конвертации.

.EXAMPLE
    .\EasyPDFConverter.ps1
    Интерактивный режим: спросит формат и DPI, обработает все PDF из папки скрипта.

.EXAMPLE
    .\EasyPDFConverter.ps1 -Format png -Dpi 600 -NoOpen
    Без вопросов: все PDF из папки скрипта в PNG при 600 DPI.

.EXAMPLE
    .\EasyPDFConverter.ps1 "D:\Сканы\договор.pdf" -Format jpeg
    Один конкретный файл в JPEG, DPI будет спрошен.

.LINK
    https://github.com/baslie/EasyPDFConverter
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,

    [ValidateSet('png', 'jpeg', 'jpg')]
    [string]$Format,

    [ValidateRange(36, 2400)]
    [int]$Dpi,

    [switch]$NoPause,

    [switch]$NoOpen
)

# ---------------------------------------------------------------------------
# Настройки окружения
# ---------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
try { [Console]::InputEncoding  = [System.Text.Encoding]::UTF8 } catch { }

$Root       = $PSScriptRoot
$ToolsDir   = Join-Path $Root 'tools'
$OutputRoot = Join-Path $Root 'output'

$PopplerRepo        = 'oschwartz10612/poppler-windows'
$PopplerFallbackTag = 'v26.02.0-0'
$PopplerFallbackUrl = "https://github.com/$PopplerRepo/releases/download/$PopplerFallbackTag/Release-26.02.0-0.zip"

# ---------------------------------------------------------------------------
# Вывод в консоль
# ---------------------------------------------------------------------------
function Write-Title {
    Write-Host ''
    Write-Host '  ==========================================' -ForegroundColor DarkCyan
    Write-Host '    EasyPDFConverter  -  PDF -> PNG / JPEG  ' -ForegroundColor Cyan
    Write-Host '  ==========================================' -ForegroundColor DarkCyan
    Write-Host ''
}
function Write-Step { param([string]$Text) Write-Host "  > $Text" -ForegroundColor Gray }
function Write-Ok   { param([string]$Text) Write-Host "  + $Text" -ForegroundColor Green }
function Write-Warn { param([string]$Text) Write-Host "  ! $Text" -ForegroundColor Yellow }
function Write-Err  { param([string]$Text) Write-Host "  x $Text" -ForegroundColor Red }

function Wait-Exit {
    param([int]$Code = 0)
    if (-not $NoPause) {
        Write-Host ''
        Write-Host '  Нажмите любую клавишу, чтобы закрыть окно...' -ForegroundColor DarkGray
        try { [void][Console]::ReadKey($true) } catch { Read-Host | Out-Null }
    }
    exit $Code
}

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N2} ГБ' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} МБ' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} КБ' -f ($Bytes / 1KB)) }
    return "$Bytes Б"
}

# ---------------------------------------------------------------------------
# Работа с путями
# ---------------------------------------------------------------------------
function Test-AsciiPath {
    param([string]$P)
    return ($P -notmatch '[^\x00-\x7F]')
}

# Возвращает папку с ASCII-путём для временных файлов, либо $null, если такой нет.
# Нужна потому, что консольные утилиты Poppler получают аргументы в системной ANSI-кодировке
# и могут не найти файл с кириллицей в имени на англоязычной Windows.
function Get-StagingDir {
    $candidates = @(
        (Join-Path $ToolsDir 'work'),
        (Join-Path $env:TEMP 'EasyPDFConverter'),
        (Join-Path $env:SystemDrive 'EasyPDFConverter.tmp')
    )
    $fso = New-Object -ComObject Scripting.FileSystemObject
    foreach ($dir in $candidates) {
        try {
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
            if (Test-AsciiPath $dir) { return $dir }
            $short = $fso.GetFolder($dir).ShortPath
            if ($short -and (Test-AsciiPath $short)) { return $short }
        } catch { }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Запуск внешних программ (без исключений PowerShell на stderr)
# ---------------------------------------------------------------------------
function Quote-Arg {
    param([string]$A)
    if ($A -match '[\s"]') { return '"' + ($A -replace '"', '\"') + '"' }
    return $A
}

function Invoke-Tool {
    param(
        [string]$Exe,
        [string[]]$Arguments,
        [string]$StdOut,
        [string]$StdErr,
        [scriptblock]$WhileRunning
    )
    $argLine = ($Arguments | ForEach-Object { Quote-Arg $_ }) -join ' '
    $psi = @{
        FilePath     = $Exe
        ArgumentList = $argLine
        NoNewWindow  = $true
        PassThru     = $true
    }
    if ($StdOut) { $psi.RedirectStandardOutput = $StdOut }
    if ($StdErr) { $psi.RedirectStandardError  = $StdErr }
    $proc = Start-Process @psi
    # Без обращения к Handle PowerShell 5.1 возвращает пустой ExitCode.
    $null = $proc.Handle
    while (-not $proc.HasExited) {
        if ($WhileRunning) { & $WhileRunning }
        Start-Sleep -Milliseconds 150
    }
    $proc.WaitForExit()
    return $proc.ExitCode
}

# ---------------------------------------------------------------------------
# Poppler: поиск и установка
# ---------------------------------------------------------------------------
function Find-Poppler {
    $local = Join-Path $ToolsDir 'poppler'
    if (Test-Path -LiteralPath $local) {
        $exe = Get-ChildItem -LiteralPath $local -Recurse -Filter 'pdftoppm.exe' -File -ErrorAction SilentlyContinue |
               Select-Object -First 1
        if ($exe) { return $exe.DirectoryName }
    }
    $cmd = Get-Command 'pdftoppm.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return (Split-Path $cmd.Source -Parent) }
    return $null
}

function Install-Poppler {
    New-Item -ItemType Directory -Force -Path $ToolsDir | Out-Null
    $popplerDir = Join-Path $ToolsDir 'poppler'
    $zip        = Join-Path $ToolsDir 'poppler.zip'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }

    $url = $null
    $tag = $null
    Write-Step 'Ищу свежую сборку Poppler на GitHub...'
    try {
        $rel   = Invoke-RestMethod -Uri "https://api.github.com/repos/$PopplerRepo/releases/latest" `
                                   -Headers @{ 'User-Agent' = 'EasyPDFConverter' } -TimeoutSec 20
        $asset = @($rel.assets | Where-Object { $_.name -like 'Release-*.zip' }) | Select-Object -First 1
        if ($asset) { $url = $asset.browser_download_url; $tag = $rel.tag_name }
    } catch {
        Write-Warn 'GitHub API недоступен, использую известную версию.'
    }
    if (-not $url) { $url = $PopplerFallbackUrl; $tag = $PopplerFallbackTag }

    Write-Step "Скачиваю Poppler $tag (около 16 МБ)..."
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add('User-Agent', 'EasyPDFConverter')
    try {
        $wc.DownloadFile($url, $zip)
    } finally {
        $wc.Dispose()
    }

    Write-Step 'Распаковываю...'
    if (Test-Path -LiteralPath $popplerDir) { Remove-Item -LiteralPath $popplerDir -Recurse -Force }
    Expand-Archive -LiteralPath $zip -DestinationPath $popplerDir -Force
    Remove-Item -LiteralPath $zip -Force

    $bin = Find-Poppler
    if (-not $bin) { throw 'В скачанном архиве не найден pdftoppm.exe.' }
    Write-Ok "Poppler готов: $bin"
    return $bin
}

function Get-PdfPageCount {
    param([string]$Bin, [string]$Pdf, [string]$Stage)
    $pdfinfo = Join-Path $Bin 'pdfinfo.exe'
    if (-not (Test-Path -LiteralPath $pdfinfo)) { return 0 }
    $tmpDir = if ($Stage) { $Stage } else { $env:TEMP }
    $out = Join-Path $tmpDir 'pdfinfo.out.txt'
    $err = Join-Path $tmpDir 'pdfinfo.err.txt'
    try {
        $code = Invoke-Tool -Exe $pdfinfo -Arguments @($Pdf) -StdOut $out -StdErr $err
        if ($code -ne 0) { return 0 }
        $m = Select-String -LiteralPath $out -Pattern '^Pages:\s+(\d+)' | Select-Object -First 1
        if ($m) { return [int]$m.Matches[0].Groups[1].Value }
    } catch { }
    finally {
        Remove-Item -LiteralPath $out, $err -Force -ErrorAction SilentlyContinue
    }
    return 0
}

# ---------------------------------------------------------------------------
# Сбор входных файлов
# ---------------------------------------------------------------------------
function Get-InputPdfs {
    param([string[]]$Items)
    $result = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    if (-not $Items -or $Items.Count -eq 0) {
        Get-ChildItem -LiteralPath $Root -Filter '*.pdf' -File | Sort-Object Name | ForEach-Object { $result.Add($_) }
        return ,$result
    }
    foreach ($item in $Items) {
        if (-not (Test-Path -LiteralPath $item)) { Write-Warn "Не найдено: $item"; continue }
        $fi = Get-Item -LiteralPath $item
        if ($fi.PSIsContainer) {
            Get-ChildItem -LiteralPath $fi.FullName -Filter '*.pdf' -File | Sort-Object Name | ForEach-Object { $result.Add($_) }
        } elseif ($fi.Extension -ieq '.pdf') {
            $result.Add($fi)
        } else {
            Write-Warn "Пропускаю (не PDF): $($fi.Name)"
        }
    }
    return ,$result
}

# ---------------------------------------------------------------------------
# Диалог с пользователем
# ---------------------------------------------------------------------------
function Read-Format {
    Write-Host '  Выберите формат:' -ForegroundColor White
    Write-Host '    [1] PNG  - без потерь, максимальное качество (рекомендуется)'
    Write-Host '    [2] JPEG - качество 100, файлы заметно меньше'
    while ($true) {
        $answer = (Read-Host '  Ваш выбор (Enter = 1)').Trim()
        switch -Regex ($answer) {
            '^(|1|png)$'  { return 'png' }
            '^(2|jpe?g)$' { return 'jpeg' }
            default       { Write-Warn 'Введите 1 или 2.' }
        }
    }
}

function Read-Dpi {
    Write-Host ''
    Write-Host '  Выберите разрешение (DPI):' -ForegroundColor White
    Write-Host '    [1] 150 - для экрана и мессенджеров'
    Write-Host '    [2] 300 - печатное качество (рекомендуется)'
    Write-Host '    [3] 600 - максимальная детализация, файлы большие'
    Write-Host '    или введите своё число от 36 до 2400'
    while ($true) {
        $answer = (Read-Host '  Ваш выбор (Enter = 2)').Trim()
        switch -Regex ($answer) {
            '^(|2)$' { return 300 }
            '^1$'    { return 150 }
            '^3$'    { return 600 }
            '^\d+$'  {
                $n = [int]$answer
                if ($n -ge 36 -and $n -le 2400) { return $n }
                Write-Warn 'Допустимый диапазон: 36-2400.'
            }
            default  { Write-Warn 'Введите 1, 2, 3 или число.' }
        }
    }
}

# ---------------------------------------------------------------------------
# Конвертация одного PDF
# ---------------------------------------------------------------------------

# Запускает pdftoppm и показывает прогресс по количеству появившихся файлов.
function Invoke-Render {
    param(
        [string]$Bin,
        [string]$InPath,
        [string]$Prefix,
        [string]$WatchDir,
        [string]$WatchPattern,
        [string]$Fmt,
        [int]$Resolution,
        [int]$Pages,
        [string]$ErrFile
    )
    $toolArgs = @('-r', "$Resolution", '-cropbox', '-aa', 'yes', '-aaVector', 'yes', '-freetype', 'yes')
    if ($Fmt -eq 'png') {
        $toolArgs += '-png'
    } else {
        $toolArgs += @('-jpeg', '-jpegopt', 'quality=100,optimize=y')
    }
    $toolArgs += @($InPath, $Prefix)

    $progress = {
        $done = @(Get-ChildItem -LiteralPath $WatchDir -Filter $WatchPattern -File -ErrorAction SilentlyContinue).Count
        if ($Pages -gt 0) {
            $line = "    страница $done из $Pages  [{0}%]" -f [int](100 * $done / $Pages)
        } else {
            $line = "    готово страниц: $done"
        }
        Write-Host ("`r" + $line.PadRight(60)) -NoNewline -ForegroundColor DarkGray
    }.GetNewClosure()

    $code = Invoke-Tool -Exe (Join-Path $Bin 'pdftoppm.exe') -Arguments $toolArgs -StdErr $ErrFile -WhileRunning $progress
    Write-Host ("`r" + ''.PadRight(60) + "`r") -NoNewline

    $errText = ''
    if (Test-Path -LiteralPath $ErrFile) {
        $raw = Get-Content -LiteralPath $ErrFile -Raw -ErrorAction SilentlyContinue
        if ($null -ne $raw) { $errText = [string]$raw }
        Remove-Item -LiteralPath $ErrFile -Force -ErrorAction SilentlyContinue
    }
    return [pscustomobject]@{ Code = $code; Err = $errText }
}

function Convert-PdfFile {
    param(
        [System.IO.FileInfo]$Pdf,
        [string]$Bin,
        [string]$Fmt,
        [int]$Resolution,
        [string]$Stage
    )
    $base   = $Pdf.BaseName
    $ext    = if ($Fmt -eq 'png') { 'png' } else { 'jpg' }
    $outDir = Join-Path $OutputRoot $base
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null

    # Удаляем результаты прошлого запуска с тем же именем, чтобы не смешивать страницы.
    Get-ChildItem -LiteralPath $outDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "$base-*.$ext" } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    $tmpDir = if ($Stage) { $Stage } else { $env:TEMP }
    $pages  = Get-PdfPageCount -Bin $Bin -Pdf $Pdf.FullName -Stage $Stage
    $pagesText = if ($pages -gt 0) { "$pages стр." } else { 'страниц: ?' }
    Write-Host ''
    Write-Host "  $($Pdf.Name)  ($(Format-Size $Pdf.Length), $pagesText)" -ForegroundColor White

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # Первая попытка - напрямую, с исходными путями.
    $r = Invoke-Render -Bin $Bin -InPath $Pdf.FullName -Prefix (Join-Path $outDir $base) `
                       -WatchDir $outDir -WatchPattern "$base-*.$ext" -Fmt $Fmt -Resolution $Resolution `
                       -Pages $pages -ErrFile (Join-Path $tmpDir 'pdftoppm.err.txt')
    $code = $r.Code
    $errText = $r.Err

    # Если не вышло и в путях есть не-ASCII символы - повторяем через временную ASCII-папку:
    # старые сборки Poppler получают аргументы в системной ANSI-кодировке и не находят такие файлы.
    $madeDirect = @(Get-ChildItem -LiteralPath $outDir -Filter "$base-*.$ext" -File -ErrorAction SilentlyContinue).Count
    $needStage  = -not ((Test-AsciiPath $Pdf.FullName) -and (Test-AsciiPath $outDir))
    if (($code -ne 0 -or $madeDirect -eq 0) -and $needStage -and $Stage) {
        Write-Host '    повторяю через временную папку с латинским путём...' -ForegroundColor DarkGray
        $stageIn  = Join-Path $Stage 'input.pdf'
        $stageOut = Join-Path $Stage 'out'
        if (Test-Path -LiteralPath $stageOut) { Remove-Item -LiteralPath $stageOut -Recurse -Force }
        New-Item -ItemType Directory -Force -Path $stageOut | Out-Null
        Copy-Item -LiteralPath $Pdf.FullName -Destination $stageIn -Force
        if ($pages -eq 0) { $pages = Get-PdfPageCount -Bin $Bin -Pdf $stageIn -Stage $Stage }

        $r = Invoke-Render -Bin $Bin -InPath $stageIn -Prefix (Join-Path $stageOut 'page') `
                           -WatchDir $stageOut -WatchPattern "page-*.$ext" -Fmt $Fmt -Resolution $Resolution `
                           -Pages $pages -ErrFile (Join-Path $Stage 'pdftoppm.err.txt')
        $code = $r.Code
        $errText = $r.Err

        Get-ChildItem -LiteralPath $stageOut -Filter "page-*.$ext" -File | ForEach-Object {
            $newName = $base + $_.Name.Substring(4)
            Move-Item -LiteralPath $_.FullName -Destination (Join-Path $outDir $newName) -Force
        }
        Remove-Item -LiteralPath $stageOut -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stageIn -Force -ErrorAction SilentlyContinue
    }
    $sw.Stop()

    $made  = @(Get-ChildItem -LiteralPath $outDir -Filter "$base-*.$ext" -File -ErrorAction SilentlyContinue)
    $total = ($made | Measure-Object -Property Length -Sum).Sum
    if ($null -eq $total) { $total = 0 }

    if ($code -ne 0 -or $made.Count -eq 0) {
        Write-Err "Не удалось сконвертировать (код $code)."
        if ($errText.Trim()) {
            ($errText.Trim() -split "`r?`n") | Select-Object -First 5 | ForEach-Object { Write-Host "      $_" -ForegroundColor DarkRed }
        }
        return [pscustomobject]@{ Ok = $false; Pages = 0; Bytes = [long]0 }
    }

    Write-Ok ("{0} стр. -> output\{1}\  ({2}, {3:N1} с)" -f $made.Count, $base, (Format-Size $total), $sw.Elapsed.TotalSeconds)
    if ($errText.Trim()) {
        Write-Warn 'Poppler сообщил о предупреждениях (файл всё равно сконвертирован).'
    }
    return [pscustomobject]@{ Ok = $true; Pages = $made.Count; Bytes = [long]$total }
}

# ---------------------------------------------------------------------------
# Основной сценарий
# ---------------------------------------------------------------------------
try {
    Write-Title

    $pdfs = Get-InputPdfs -Items $Path
    if ($pdfs.Count -eq 0) {
        Write-Warn 'PDF-файлы не найдены.'
        Write-Host ''
        Write-Host '  Положите PDF-файлы в эту папку:' -ForegroundColor White
        Write-Host "    $Root"
        Write-Host '  и запустите скрипт ещё раз. Или перетащите PDF прямо на EasyPDFConverter.bat.'
        Wait-Exit 1
    }

    Write-Host "  Найдено PDF: $($pdfs.Count)" -ForegroundColor White
    foreach ($p in $pdfs) { Write-Host ("    - {0}  ({1})" -f $p.Name, (Format-Size $p.Length)) }
    Write-Host ''

    if (-not $Format) { $Format = Read-Format } else { Write-Step "Формат: $Format" }
    if ($Format -eq 'jpg') { $Format = 'jpeg' }
    if (-not $Dpi)    { $Dpi = Read-Dpi }       else { Write-Step "DPI: $Dpi" }
    Write-Host ''

    $bin = Find-Poppler
    if (-not $bin) {
        Write-Step 'Poppler ещё не установлен - скачиваю портативную сборку (один раз).'
        $bin = Install-Poppler
    } else {
        Write-Step "Poppler: $bin"
    }

    $stage = Get-StagingDir
    New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null

    $fmtLabel = if ($Format -eq 'png') { 'PNG (без потерь)' } else { 'JPEG (качество 100)' }
    Write-Host ''
    Write-Host "  Конвертирую в $fmtLabel, $Dpi DPI" -ForegroundColor Cyan

    $okFiles = 0; $failFiles = 0; $totalPages = 0; $totalBytes = [long]0
    $swAll = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($pdf in $pdfs) {
        try {
            $r = Convert-PdfFile -Pdf $pdf -Bin $bin -Fmt $Format -Resolution $Dpi -Stage $stage
            if ($r.Ok) { $okFiles++; $totalPages += $r.Pages; $totalBytes += $r.Bytes } else { $failFiles++ }
        } catch {
            $failFiles++
            Write-Err "Ошибка при обработке $($pdf.Name): $($_.Exception.Message)"
        }
    }
    $swAll.Stop()

    Write-Host ''
    Write-Host '  ------------------------------------------' -ForegroundColor DarkCyan
    if ($okFiles -gt 0) {
        Write-Ok ("Готово: {0} файл(ов), {1} страниц, {2} за {3:N1} с" -f $okFiles, $totalPages, (Format-Size $totalBytes), $swAll.Elapsed.TotalSeconds)
        Write-Host "  Результат: $OutputRoot" -ForegroundColor White
    }
    if ($failFiles -gt 0) { Write-Err "С ошибками: $failFiles файл(ов)." }

    if ($okFiles -gt 0 -and -not $NoOpen) {
        try { Invoke-Item -LiteralPath $OutputRoot } catch { }
    }

    $exitCode = 0
    if ($failFiles -gt 0 -and $okFiles -eq 0) { $exitCode = 1 }
    Wait-Exit $exitCode
}
catch {
    Write-Host ''
    Write-Err "Критическая ошибка: $($_.Exception.Message)"
    if ($_.ScriptStackTrace) {
        ($_.ScriptStackTrace -split "`n") | Select-Object -First 3 | ForEach-Object { Write-Host "      $_" -ForegroundColor DarkRed }
    }
    Wait-Exit 1
}
