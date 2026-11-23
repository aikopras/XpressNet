@echo off
rem XpressNet Qt examples: build all examples for Windows (November 2026)
rem Start with a double click, or from a command prompt in this directory. See Examples-Windows.md.
rem Builds the library and the six examples (Release), and makes for every example a zip file with the program
rem and the Qt DLLs in ..\..\compiled\Windows-Qt. The directories build and deploy are first removed.
rem Requires Visual Studio (with "C++ CMake tools for Windows") and Qt 6 for MSVC (with Qt Serial Port).

setlocal

rem ---- Settings: adapt these three lines to the installed Visual Studio and Qt versions --------------------------
rem Visual Studio 2022: GENERATOR=Visual Studio 17 2022, VSVERSION=[17.0,18.0), Qt 6.8 or later with msvc2022_64
set "QTDIR=C:\Qt\6.7.3\msvc2019_64"
set "GENERATOR=Visual Studio 16 2019"
set "VSVERSION=[16.0,17.0)"
rem -------------------------------------------------------------------------------------------------------------

cd /d "%~dp0"

rem The library source must be found relative to this directory (the complete repository must be unzipped)
if not exist "..\..\..\Library\src\Qt\XpressNetQt\CMakeLists.txt" goto nolibrary

rem Qt
if not exist "%QTDIR%\bin\windeployqt.exe" goto noqt

rem Visual Studio and its CMake, found with vswhere (installed together with Visual Studio)
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" goto novs
set "VSPATH="
"%VSWHERE%" -version "%VSVERSION%" -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath > "%TEMP%\xpn-vspath.txt"
set /p VSPATH=<"%TEMP%\xpn-vspath.txt"
del "%TEMP%\xpn-vspath.txt"
if not defined VSPATH goto novs
set "CMAKE=%VSPATH%\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe"
if not exist "%CMAKE%" goto nocmake

echo Qt:            %QTDIR%
echo Visual Studio: %VSPATH%
echo.

rem Build the library and all examples
if exist build rmdir /s /q build
if exist deploy rmdir /s /q deploy
"%CMAKE%" -S . -B build -G "%GENERATOR%" -A x64 -DCMAKE_PREFIX_PATH="%QTDIR%"
if errorlevel 1 goto error
"%CMAKE%" --build build --config Release
if errorlevel 1 goto error

rem One zip file per example
if not exist "..\..\compiled\Windows-Qt" mkdir "..\..\compiled\Windows-Qt"
call :deploy "XpressNet Basic Test"
if errorlevel 1 goto error
call :deploy "XpressNet Test"
if errorlevel 1 goto error
call :deploy "Z21 Basic Test"
if errorlevel 1 goto error
call :deploy "XpressNet Multiple Instances"
if errorlevel 1 goto error
call :deploy "RS-Bus Monitor"
if errorlevel 1 goto error
call :deploy "RS-Bus Decoder Scanner"
if errorlevel 1 goto error

echo.
echo Ready. The zip files are in Examples\compiled\Windows-Qt, the programs with their Qt DLLs in deploy.
pause
exit /b 0

rem Copies the program of one example (%1) to deploy\<example>, adds the Qt DLLs and makes the zip file
:deploy
echo.
echo ==== %~1
mkdir "deploy\%~1"
copy /y "build\%~1\Release\%~1.exe" "deploy\%~1" >nul
if errorlevel 1 exit /b 1
"%QTDIR%\bin\windeployqt.exe" --release --no-translations --no-opengl-sw --no-system-d3d-compiler --no-compiler-runtime "deploy\%~1\%~1.exe"
if errorlevel 1 exit /b 1
rem The DirectX shader compilers that windeployqt may add are not used by the programs
if exist "deploy\%~1\dxcompiler.dll" del "deploy\%~1\dxcompiler.dll"
if exist "deploy\%~1\dxil.dll" del "deploy\%~1\dxil.dll"
if exist "..\..\compiled\Windows-Qt\%~1.zip" del "..\..\compiled\Windows-Qt\%~1.zip"
tar -a -c -f "..\..\compiled\Windows-Qt\%~1.zip" -C deploy "%~1"
if errorlevel 1 exit /b 1
exit /b 0

:nolibrary
echo ERROR: ..\..\..\Library\src\Qt\XpressNetQt not found.
echo Unzip the complete repository and start this file in its directory Examples\src\Qt.
goto stop

:noqt
echo ERROR: %QTDIR%\bin\windeployqt.exe not found.
echo Install Qt for MSVC, or change QTDIR at the top of this file.
goto stop

:novs
echo ERROR: Visual Studio %VSVERSION% with C++ not found.
echo Install Visual Studio with the workload "Desktop development with C++",
echo or change GENERATOR and VSVERSION at the top of this file.
goto stop

:nocmake
echo ERROR: CMake not found in %VSPATH%.
echo Add the component "C++ CMake tools for Windows" with the Visual Studio Installer.
goto stop

:error
echo.
echo ERROR: the build stopped, see the messages above.

:stop
pause
exit /b 1
