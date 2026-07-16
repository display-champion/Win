@echo off
chcp 65001 >nul
title オンラインで協力 一覧 - 価格のみ高速更新
echo ================================================================
echo  価格のみ高速更新 (一覧と価格を更新。過去最安は前回値を流用)
echo ================================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1" -SkipDeku
echo.
pause
