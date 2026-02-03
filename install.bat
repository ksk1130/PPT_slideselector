@echo off
REM ===================================================================
REM PowerPoint スライド選択 アドイン インストールスクリプト
REM ===================================================================

REM 日本語対応
chcp 65001 > nul

echo.
echo ===================================================================
echo PowerPoint スライド選択 アドイン インストール
echo ===================================================================
echo.

REM AddinFolder のパスを定義
set "AddinFolder=%APPDATA%\Microsoft\AddIns"

REM AddinFolder が存在しない場合は作成
if not exist "%AddinFolder%" (
    echo AdIns フォルダを作成しています...
    mkdir "%AddinFolder%"
)

REM スクリプトのあるディレクトリを取得
set "ScriptDir=%~dp0"

REM コピー元ファイル
set "SourceFile=%ScriptDir%slideselector.pptm"

REM コピー先ファイル
set "DestFile=%AddinFolder%\slideselector.pptm"

REM ファイルの存在確認
if not exist "%SourceFile%" (
    echo.
    echo エラー：slideselector.pptm が見つかりません。
    echo このスクリプトと同じフォルダに slideselector.pptm を配置してください。
    echo.
    pause
    exit /b 1
)

REM 既存ファイルがあれば削除
if exist "%DestFile%" (
    echo 既存のアドインを削除しています...
    del "%DestFile%"
)

REM ファイルをコピー
echo アドインをインストールしています...
copy "%SourceFile%" "%DestFile%"

if errorlevel 1 (
    echo.
    echo エラー：インストール中にエラーが発生しました。
    echo 管理者権限で実行してください。
    echo.
    pause
    exit /b 1
)

echo.
echo ===================================================================
echo インストール完了しました！
echo ===================================================================
echo.
echo 次のステップ：
echo 1. PowerPoint を完全に終了してください（全てのウィンドウを閉じる）
echo 2. PowerPoint を再起動してください
echo 3. 「ホーム」タブに「スライド操作」グループが表示されます
echo 4. 「スライド選択」ボタンをクリックしてご使用ください
echo.
pause
