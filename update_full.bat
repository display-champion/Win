@echo off
chcp 65001 >nul
title オンラインで協力 一覧 - フル更新
echo ================================================================
echo  フル更新 (一覧 + 過去最安を取り直し)  ※数分〜十数分かかります
echo ================================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1"
echo.
pause
