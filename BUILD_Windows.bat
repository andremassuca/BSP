@echo off
chcp 65001 >nul 2>&1
title BSP - Biomechanical Stability Program - Build Installer

echo.
echo ============================================================
echo   BSP - Biomechanical Stability Program  v1.0
echo   Building BSP_Setup.exe  (standalone installer)
echo ============================================================
echo.

:: Find Python
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if "%PYTHON%"=="" where python3 >nul 2>&1 && set PYTHON=python3
if "%PYTHON%"=="" (
    if exist "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" set PYTHON=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe
    if exist "%LOCALAPPDATA%\Python\pythoncore-3.13-64\python.exe" set PYTHON=%LOCALAPPDATA%\Python\pythoncore-3.13-64\python.exe
    if exist "%LOCALAPPDATA%\Python\pythoncore-3.12-64\python.exe" set PYTHON=%LOCALAPPDATA%\Python\pythoncore-3.12-64\python.exe
    if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" set PYTHON=%LOCALAPPDATA%\Programs\Python\Python314\python.exe
    if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
)
if "%PYTHON%"=="" (
    echo   [ERROR] Python not found.
    echo   Install from: https://www.python.org/downloads/
    echo   Enable "Add Python to PATH" during installation.
    pause & exit /b 1
)
echo   Python: %PYTHON%
%PYTHON% --version
echo.

:: Ensure icon has the correct name
if exist AOM.ico (
    if not exist BSP.ico copy /y AOM.ico BSP.ico >nul
)

:: Run tests before building
echo   [0/4]  Running test suite...
%PYTHON% -X utf8 estabilidade_gui.py --testes >nul 2>&1
if errorlevel 1 (
    echo   [WARNING] Some tests failed - proceeding anyway
) else (
    echo          Tests OK
)
echo.

:: Dependencies
echo   [1/4]  Installing dependencies...
%PYTHON% -m pip install pyinstaller nuitka "nuitka[onefile]" --upgrade --quiet
if exist requirements.txt (
    %PYTHON% -m pip install -r requirements.txt --upgrade --quiet
) else (
    %PYTHON% -m pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow --upgrade --quiet
)
if errorlevel 1 ( echo   [ERROR] pip failed. & pause & exit /b 1 )
echo          OK
echo.

:: Step 2: BSP.exe (main application)
:: Nuitka compila Python para codigo maquina nativo (C) - muito mais dificil
:: de reverter do que bytecode PyInstaller. Nao requer licenca.
echo   [2/4]  Compiling BSP.exe with Nuitka  (15-30 min first time)...
echo          (compilacao nativa: codigo maquina, impossivel descompilar)
echo.
if exist dist rmdir /s /q dist
mkdir dist
%PYTHON% -m nuitka --onefile --windows-console-mode=disable ^
    --windows-icon-from-ico=BSP.ico ^
    --output-filename=dist\BSP.exe ^
    --enable-plugin=tk-inter ^
    --enable-plugin=numpy ^
    --assume-yes-for-downloads ^
    --jobs=4 ^
    --quiet ^
    estabilidade_gui.py
if errorlevel 1 ( echo   [ERROR] BSP.exe compilation failed. & pause & exit /b 1 )
echo          dist\BSP.exe OK
echo.

:: Step 3: BSP_Uninstall.exe (standalone uninstaller)
echo   [3/4]  Compiling BSP_Uninstall.exe...
echo.
%PYTHON% -m PyInstaller --onefile --windowed --name BSP_Uninstall --icon BSP.ico ^
    --hidden-import winreg --hidden-import tkinter --hidden-import tkinter.ttk ^
    bsp_uninstaller.py
if errorlevel 1 (
    echo   [WARNING] Uninstaller GUI did not compile - fallback .bat will be used
) else (
    echo          dist\BSP_Uninstall.exe OK
)
echo.

:: Step 4: BSP_Setup.exe (bundles BSP.exe + BSP_Uninstall.exe + icon)
echo   [4/4]  Compiling BSP_Setup.exe...
echo.
if exist dist\BSP_Uninstall.exe (
    %PYTHON% -m PyInstaller --onefile --windowed --name BSP_Setup --icon BSP.ico ^
        --add-data "dist\BSP.exe;." --add-data "dist\BSP_Uninstall.exe;." --add-data "BSP.ico;." ^
        --hidden-import winreg --hidden-import tkinter --hidden-import tkinter.ttk ^
        --hidden-import tkinter.filedialog --hidden-import PIL ^
        --hidden-import PIL.Image --hidden-import PIL.ImageTk ^
        bsp_installer.py
) else (
    %PYTHON% -m PyInstaller --onefile --windowed --name BSP_Setup --icon BSP.ico ^
        --add-data "dist\BSP.exe;." --add-data "BSP.ico;." ^
        --hidden-import winreg --hidden-import tkinter --hidden-import tkinter.ttk ^
        --hidden-import tkinter.filedialog --hidden-import PIL ^
        --hidden-import PIL.Image --hidden-import PIL.ImageTk ^
        bsp_installer.py
)
if errorlevel 1 ( echo   [ERROR] BSP_Setup.exe compilation failed. & pause & exit /b 1 )

copy /y dist\BSP_Setup.exe BSP_Setup.exe >nul
echo.
echo ============================================================
echo   BUILD COMPLETE!
echo.
echo   BSP_Setup.exe  -  distribute this file
echo.
echo   Installation creates:
echo     - Desktop shortcut with BSP icon
echo     - Start Menu entry
echo     - BSP_Uninstall.exe with its own GUI
echo     - Entry in Add/Remove Programs
echo       (Uninstall opens the correct BSP_Uninstall.exe)
echo ============================================================
echo.
pause
