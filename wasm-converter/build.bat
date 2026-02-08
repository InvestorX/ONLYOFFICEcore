@echo off
REM build.bat - WASMビルドスクリプト（Windows用）
REM
REM 使用方法:
REM   build.bat          # リリースビルド
REM   build.bat dev      # デバッグビルド
REM   build.bat test     # テスト実行
REM   build.bat serve    # ローカルサーバー起動
REM   build.bat fonts    # フォントダウンロード

setlocal enabledelayedexpansion

REM スクリプトのディレクトリに移動
cd /d "%~dp0"

REM コマンドライン引数の処理
set "MODE=%~1"
if "%MODE%"=="" set "MODE=build"

REM メイン処理
if /i "%MODE%"=="build" goto :build_release
if /i "%MODE%"=="release" goto :build_release
if /i "%MODE%"=="dev" goto :build_dev
if /i "%MODE%"=="test" goto :run_tests
if /i "%MODE%"=="serve" goto :serve
if /i "%MODE%"=="fonts" goto :download_fonts

REM 無効なコマンドの場合
echo 使用方法: %~nx0 {build^|dev^|test^|serve^|fonts}
echo.
echo コマンド:
echo   build   - リリースモードでWASMをビルド（デフォルト）
echo   dev     - デバッグモードでWASMをビルド
echo   test    - ネイティブテストを実行
echo   serve   - ローカルHTTPサーバーを起動
echo   fonts   - 日本語フォントをダウンロード
exit /b 1

:build_release
call :download_fonts
call :build_wasm release
exit /b 0

:build_dev
call :download_fonts
call :build_wasm dev
exit /b 0

:run_tests
call :info "テストを実行中..."
cargo test --lib -- --nocapture
if errorlevel 1 (
    call :error "テストが失敗しました"
    exit /b 1
)
call :success "全テスト合格！"
exit /b 0

:serve
call :info "ローカルサーバーを起動中 (http://localhost:8080)..."
cd www
python -m http.server 8080 2>nul
if errorlevel 1 (
    python -m SimpleHTTPServer 8080 2>nul
    if errorlevel 1 (
        call :error "Python が見つかりません。別の HTTP サーバーを使用してください。"
        exit /b 1
    )
)
exit /b 0

:download_fonts
if exist "fonts\NotoSansJP-Regular.ttf" (
    goto :eof
)
call :info "日本語フォント (Noto Sans JP) をダウンロード中..."
if not exist "fonts" mkdir fonts

set "FONT_URL=https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf"

REM curlを試す
where curl >nul 2>nul
if %errorlevel%==0 (
    curl -L -o "fonts\NotoSansJP-Regular.ttf" "%FONT_URL%" 2>nul
    if exist "fonts\NotoSansJP-Regular.ttf" goto :eof
)

REM PowerShellを試す
where powershell >nul 2>nul
if %errorlevel%==0 (
    powershell -Command "try { Invoke-WebRequest -Uri '%FONT_URL%' -OutFile 'fonts\NotoSansJP-Regular.ttf' } catch { exit 1 }" 2>nul
    if exist "fonts\NotoSansJP-Regular.ttf" goto :eof
)

call :info "フォントのダウンロードに失敗しました。embed-fontsフィーチャーなしでビルドします。"
goto :eof

:build_wasm
set "BUILD_MODE=%~1"
call :check_wasm_pack
if errorlevel 1 exit /b 1

if /i "%BUILD_MODE%"=="dev" (
    call :info "デバッグモードでWASMをビルド中..."
    wasm-pack build --target web --dev --out-dir www/pkg
) else (
    call :info "リリースモードでWASMをビルド中..."
    wasm-pack build --target web --release --out-dir www/pkg
)

if errorlevel 1 (
    call :error "WASMビルドが失敗しました"
    exit /b 1
)

call :success "WASMビルド完了！ www/pkg/ に出力されました。"
goto :eof

:check_wasm_pack
where wasm-pack >nul 2>nul
if %errorlevel%==0 goto :eof

call :info "wasm-pack をインストール中..."
cargo install wasm-pack
if errorlevel 1 (
    call :error "wasm-packのインストールに失敗しました"
    exit /b 1
)
goto :eof

:info
echo [INFO] %~1
goto :eof

:success
echo [OK] %~1
goto :eof

:error
echo [ERROR] %~1
goto :eof
