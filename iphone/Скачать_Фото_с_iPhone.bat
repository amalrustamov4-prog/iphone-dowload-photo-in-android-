@echo off
chcp 65001 >nul
title Перенос фото с iPhone на компьютер
echo ================================================================
echo Запуск копирования фото и видео с iPhone...
echo ================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0TransferPhotos.ps1"

echo.
echo Нажмите любую клавишу для выхода...
pause >nul
