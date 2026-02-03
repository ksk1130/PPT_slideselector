@echo off
REM ===================================================================
REM PowerPoint スライド選択 アドイン アンインストールスクリプト
REM ===================================================================

REM 日本語対応
chcp 65001 > nul

echo.
echo ===================================================================
echo PowerPoint スライド選択 アドイン アンインストール
echo ===================================================================
echo.

REM AddinFolder のパスを定義
set "AddinFolder=%APPDATA%\Microsoft\AddIns"

REM アンインストール対象ファイル
set "TargetFile=%AddinFolder%\SlideJumper.pptm"

REM ファイルの存在確認
if not exist "%TargetFile%" (
    echo.
    echo アドインがインストールされていません。
    echo.
    pause
    exit /b 0
)

echo アドインをアンインストールしています...

REM ファイルを削除
del "%TargetFile%"

if errorlevel 1 (
    echo.
    echo エラー：アンインストール中にエラーが発生しました。
    echo PowerPoint が起動している場合は、終了してから再実行してください。
    echo.
    pause
    exit /b 1
)

echo.
echo ===================================================================
echo アンインストール完了しました
echo ===================================================================
echo.
echo PowerPoint を再起動すると、アドインが削除されます。
echo.
pause
