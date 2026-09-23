# ==============================================================================
# Скрипт автоматического переноса всех фото и видео с iPhone на компьютер через USB
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "       ПЕРЕНОС ВСЕХ ФОТО И ВИДЕО С IPHONE НА КОМПЬЮТЕР (USB)    " -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Папка назначения
$targetDir = Join-Path $PSScriptRoot "iPhone_Photos"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

Write-Host "[i] Фото и видео будут сохранены в:" -ForegroundColor Cyan
Write-Host "    $targetDir" -ForegroundColor White
Write-Host ""

# Шаг 1: Поиск подключенного iPhone
Write-Host "[1/3] Поиск iPhone, подключенного по USB..." -ForegroundColor Yellow
$iphoneItem = $null
$shell = $null

while ($true) {
    $shell = New-Object -ComObject Shell.Application
    $thisPC = $shell.Namespace(17) # 17 = ssfDRIVES ("Этот компьютер")
    
    foreach ($item in $thisPC.Items()) {
        if ($item.Name -match "iPhone|iPad|Apple") {
            $iphoneItem = $item
            break
        }
    }
    
    if ($iphoneItem) {
        Write-Host "  [OK] Найден: $($iphoneItem.Name)!" -ForegroundColor Green
        break
    }
    
    Write-Host "  [!] iPhone не найден. Подключите iPhone кабелем USB к компьютеру..." -ForegroundColor DarkYellow
    Start-Sleep -Seconds 2
}

# Шаг 2: Проверка разблокировки экрана (доступ к DCIM)
Write-Host ""
Write-Host "[2/3] Проверка доступа к памяти iPhone..." -ForegroundColor Yellow

$dcimItem = $null
$warned = $false

while ($true) {
    # Свежий объект Shell для сброса кэша проводника
    $shell = New-Object -ComObject Shell.Application
    $thisPC = $shell.Namespace(17)
    $iphoneItem = $thisPC.Items() | Where-Object { $_.Name -match "iPhone|iPad|Apple" } | Select-Object -First 1
    
    if (-not $iphoneItem) {
        Write-Host "  [!] Связь с iPhone потеряна. Проверьте USB-кабель." -ForegroundColor Red
        Start-Sleep -Seconds 2
        continue
    }

    $iphoneFolder = $iphoneItem.GetFolder
    if ($iphoneFolder) {
        $storage = $iphoneFolder.Items() | Where-Object { $_.Name -match "Storage|Память|Хранилище|Internal" } | Select-Object -First 1
        if (-not $storage) {
            $storage = $iphoneFolder.Items() | Select-Object -First 1
        }

        if ($storage) {
            $storageFolder = $storage.GetFolder
            if ($storageFolder -and $storageFolder.Items().Count -gt 0) {
                $dcimItem = $storageFolder.Items() | Where-Object { $_.Name -match "DCIM" } | Select-Object -First 1
                if ($dcimItem) {
                    Write-Host "  [OK] Доступ к фото получен (папка DCIM открыта)!" -ForegroundColor Green
                    break
                }
            }
        }
    }

    if (-not $warned) {
        Write-Host ""
        Write-Host "  --------------------------------------------------------" -ForegroundColor DarkCyan
        Write-Host "  ВАЖНО! ПАМЯТЬ IPHONE ЗАБЛОКИРОВАНА." -ForegroundColor Red
        Write-Host "  1. Разблокируйте экран вашего iPhone (введите код-пароль)." -ForegroundColor Yellow
        Write-Host "  2. Если появится вопрос 'Доверять этому компьютеру?'" -ForegroundColor Yellow
        Write-Host "     нажмите 'ДОВЕРЯТЬ'." -ForegroundColor Green
        Write-Host "  --------------------------------------------------------" -ForegroundColor DarkCyan
        Write-Host "  Ожидание разблокировки экрана..." -ForegroundColor Gray
        $warned = $true
    } else {
        Write-Host "  ..." -ForegroundColor Gray
    }

    Start-Sleep -Seconds 2
}

# Шаг 3: Сканирование файлов в DCIM
Write-Host ""
Write-Host "[3/3] Сканирование всех фото и видео на телефоне..." -ForegroundColor Yellow

function Get-AllFilesRecursive($folderObj) {
    $fileList = [System.Collections.Generic.List[Object]]::new()
    $folder = $folderObj.GetFolder
    if (-not $folder) { return $fileList }

    foreach ($child in $folder.Items()) {
        if ($child.IsFolder) {
            $subFiles = Get-AllFilesRecursive $child
            foreach ($sf in $subFiles) { $fileList.Add($sf) }
        } else {
            $fileList.Add($child)
        }
    }
    return $fileList
}

$allMedia = Get-AllFilesRecursive $dcimItem
$totalCount = $allMedia.Count

Write-Host "  Всего найдено файлов на iPhone: $totalCount" -ForegroundColor Cyan
if ($totalCount -eq 0) {
    Write-Host "  В галерее нет фотографий для скачивания." -ForegroundColor DarkYellow
    Read-Host "Нажмите Enter для завершения..."
    exit
}

# Шаг 4: Копирование файлов
Write-Host ""
Write-Host "Начинаем копирование..." -ForegroundColor Green
$destShellFolder = $shell.Namespace($targetDir)

$copiedCount = 0
$skippedCount = 0
$errorCount = 0
$index = 0
$startTime = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($item in $allMedia) {
    $index++
    $fileName = $item.Name
    $destFile = Join-Path $targetDir $fileName
    $percent = [math]::Round(($index / $totalCount) * 100, 1)

    # Проверяем, был ли файл уже скачан ранее
    if (Test-Path $destFile) {
        $existing = Get-Item $destFile -ErrorAction SilentlyContinue
        if ($existing -and $existing.Length -gt 0) {
            $skippedCount++
            Write-Host "[$index/$totalCount - $percent%] [ПРОПУСК] $fileName (уже скачан)" -ForegroundColor DarkGray
            continue
        }
    }

    Write-Host "[$index/$totalCount - $percent%] Копирование: $fileName..." -ForegroundColor White -NoNewline

    try {
        # 16 = Respond with "Yes to All" for any dialog
        $destShellFolder.CopyHere($item, 16)

        # Ожидание завершения копирования файла
        $waitWatch = [System.Diagnostics.Stopwatch]::StartNew()
        $maxWaitSeconds = 180 # до 3 минут для больших 4K видео
        $copySuccess = $false

        while (-not (Test-Path $destFile) -and ($waitWatch.Elapsed.TotalSeconds -lt $maxWaitSeconds)) {
            Start-Sleep -Milliseconds 100
        }

        while ($waitWatch.Elapsed.TotalSeconds -lt $maxWaitSeconds) {
            try {
                $fileStream = [System.IO.File]::Open($destFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
                $fileStream.Close()
                $fileStream.Dispose()
                $copySuccess = $true
                break
            } catch {
                Start-Sleep -Milliseconds 150
            }
        }

        if ($copySuccess) {
            $copiedCount++
            $sizeMB = [math]::Round((Get-Item $destFile).Length / 1MB, 2)
            Write-Host " [ГОТОВО] ($sizeMB МБ)" -ForegroundColor Green
        } else {
            $errorCount++
            Write-Host " [ТАЙМАУТ]" -ForegroundColor Red
        }
    } catch {
        $errorCount++
        Write-Host " [ОШИБКА: $($_.Exception.Message)]" -ForegroundColor Red
    }
}

$startTime.Stop()
$elapsedMinutes = [math]::Round($startTime.Elapsed.TotalMinutes, 1)

# Итоги
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "                      КОПИРОВАНИЕ ЗАВЕРШЕНО!                   " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  - Новых файлов скопировано: $copiedCount" -ForegroundColor Green
Write-Host "  - Пропущено (были скачаны ранее): $skippedCount" -ForegroundColor DarkGray
if ($errorCount -gt 0) {
    Write-Host "  - Ошибок: $errorCount" -ForegroundColor Red
}
Write-Host "  - Затрачено времени: $elapsedMinutes мин." -ForegroundColor White
Write-Host "  - Папка с файлами: $targetDir" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Открываем папку с сохраненными фото в проводнике
Invoke-Item $targetDir

Write-Host "Папка с фотографиями открыта в проводнике." -ForegroundColor Green
Write-Host ""
