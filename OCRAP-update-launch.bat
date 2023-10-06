@echo off
setlocal enableextensions enabledelayedexpansion

echo OCRAP-VA auto updating launcher
echo This script updates and launches the Oracle Cerner Repository for Automating Pharmacy-VA (OCRAP-VA) by Lewis DeJaegher
echo.

REM Define the URL of your AutoHotkey script and xml meta files
set "repo_base_url=https://raw.githubusercontent.com/PBMCI/AHK-OCRAP/main"
set "script_url=%repo_base_url%/Oracle_Cerner_Repository_for_Automating_Pharmacy-VA.ahk"
set "meta_url=%repo_base_url%/ocrap_meta.xml"

REM Define the local path where the meta file is
set "target_directory=%USERPROFILE%\AppData\Roaming\OCRAP"
set "local_meta_file=%target_directory%\ocrap_meta.xml"
set "local_script_file=%target_directory%\ocrap-va.ahk"
set "temp_meta_file=%target_directory%\temp_meta.xml"

REM Reset errorlevel to 0
set errorlevel=0

REM Check if the directory exists
if not exist "%target_directory%" (
    echo This must be new for you...
    echo Creating your OCRAP directory here: %target_directory%
    mkdir "%target_directory%"
    echo.
)

REM Check if the script file and meta file exist locally, download if not present
if not exist "%local_meta_file%" (
    goto DownloadFiles
)

if not exist "%local_script_file%" (
    goto DownloadFiles
)

REM If we got here naturally, skip download.
goto Pass1

:DownloadFiles
echo Downloading OCRAP-VA.ahk
curl -o "%local_script_file%" "%script_url%" --ssl-no-revoke -s || (
    echo Error downloading ahk file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    exit /b 1
)

REM Reset errorlevel to 0
set errorlevel=0

echo.

REM Check the content of the downloaded AHK file for "404: Not Found" - github download issue fix
type "%local_script_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    echo 404 Error occurred while downloading ahk script file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    exit /b 1
)

echo Downloading ocrap_meta.xml
curl -o "%local_meta_file%" "%meta_url%" --ssl-no-revoke -s || (
    echo Error downloading meta file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    exit /b 1
)

REM Reset errorlevel to 0
set errorlevel=0

echo.

REM Check the content of the downloaded Meta file for "404: Not Found" - github download issue fix
type "%local_meta_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    echo 404 Error occurred while downloading metadata file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    echo This is just metadata to assist with updates; the script appears to be available.
    goto retryPrompt
)

REM Other situations wouldn't apply, go straight to launch
goto LaunchScript

:Pass1

REM If the file is found, check the current version
echo The launcher is comparing your local version against the latest build.
REM get remote version
echo Downloading build info...
curl -o "%temp_meta_file%" "%meta_url%" --ssl-no-revoke -s || (
    echo Error downloading Meta file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    goto retryPrompt
)

REM Reset errorlevel to 0
set errorlevel=0

REM Check the content of the downloaded Meta file for "404: Not Found" - github download issue fix
type "%temp_meta_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    echo 404 Error occurred while downloading metadata file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    goto retryPrompt
)

set "buildversion="
for /f "tokens=3 delims=<>" %%a in (
    'find /i "<Version>" ^< "%temp_meta_file%"'
) do set "buildversion=%%a"

echo The current OCRAP-VA.ahk build is %buildversion%

REM get local version
set "installedversion="
for /f "tokens=3 delims=<>" %%a in (
    'find /i "<Version>" ^< "%local_meta_file%"'
) do set "installedversion=%%a"

echo Your local OCRAP-VA.ahk version is %installedversion%
echo.

REM compare local to remote version

if "%installedversion%" EQU "%buildversion%" (
    REM if versions are the same, run the local script, you're done!
    echo You are up to date.
    del "%temp_meta_file%"
    goto LaunchScript
)

REM if versions are NOT the same, download remote before running
echo The local version is different from the build version. The latest build will now be downloaded.
curl -o "%local_script_file%" "%script_url%" --ssl-no-revoke -s || (
    echo Error downloading ahk file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    echo.
    echo Attempting to start the old version of OCRAP-VA.ahk
    pause
    goto LaunchScript
)

REM Reset errorlevel to 0
set errorlevel=0

REM Check the content of the downloaded file for "404: Not Found" - github download issue fix
type "%local_script_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    echo 404 Error occurred while downloading script file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    exit /b 1
)

echo.
echo Success. You are now on the latest version.
REM Copy the contents of temp_meta_file to local_meta_file, overwriting it
copy /y "%temp_meta_file%" "%local_meta_file%"

REM Check if the copy was successful
if %errorlevel% equ 0 (
    REM Delete temp_meta_file
    del "%temp_meta_file%"
) else (
    echo.
    echo Error updating local metadata file.
)

REM in any case, start the script
:LaunchScript
REM Check if AutoHotkey is installed - need to verify path
set "ahk_exe_path=C:\Program Files\AutoHotkey\v2\AutoHotkey.exe"

if not exist "%ahk_exe_path%" (
    echo.
    echo Error! - AutoHotkey v2 does not seem to be installed on this device.
    echo Autohotkey v2 is required for this script to work. This program is TRM approved.
    echo See the related Pharmacy EHRM Community of Practice Channel thread for more info.
    echo.

    REM Prompt the user for input
:retryPrompt
    set "user_choice="
    set /p "user_choice=Would you like to try to start the script anyways? Y or N: "
    REM Check the user's choice and determine the next action
    if /i "!user_choice!"=="y" (
        echo.
        echo Attempting script start...
    ) else if /i "!user_choice!"=="n" (
        echo.
        echo Exiting
        timeout /t 3
        exit /b 1
    ) else (
        echo Invalid choice.
        echo Please enter any of Y, N, y, n.
        goto retryPrompt
        pause
    )
)

start "" "%local_script_file%"

REM Check if the run was successful
if %errorlevel% equ 0 (
    echo OCRAP-VA.ahk has Launched.
    echo.
    timeout /t 5
) else (
    echo Launch Attempted despite errors.
    echo Check your windows taskbar tray to confirm whether the ahk script is running.
    echo.
    pause
)
