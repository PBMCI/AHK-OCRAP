@echo off
setlocal enableextensions enabledelayedexpansion

echo OCRAP-VA auto updating launcher
echo This script updates and launches the Oracle Cerner Repository for Automating Pharmacy-VA (OCRAP-VA) by Lewis DeJaegher
echo.

:: Define the URL of your AutoHotkey script and xml meta files
set "repo_base_url=https://raw.githubusercontent.com/PBMCI/AHK-OCRAP/main"
set "script_url=%repo_base_url%/Oracle_Cerner_Repository_for_Automating_Pharmacy-VA.ahk"
set "meta_url=%repo_base_url%/ocrap_meta.xml"
set "icon_url=%repo_base_url%/images/ocrap.ico"
set "launcher_url=%repo_base_url%/OCRAP-update-launch.bat"

:: Define the local path where the meta file is
set "target_directory=%USERPROFILE%\AppData\Roaming\OCRAP"
set "local_meta_file=%target_directory%\ocrap_meta.xml"
set "local_script_file=%target_directory%\ocrap-va.ahk"
set "temp_meta_file=%target_directory%\temp_meta.xml"
set "icon_file=%target_directory%\ocrap.ico"
set "launcher_file=%target_directory%\OCRAP-update-launch.bat"
set "startup_folder=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
for /f "delims=" %%i in ('powershell -C "[Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)"') do set "desktop_folder=%%i"

:: Reset errorlevel to 0
set errorlevel=0

:: Check if the directory exists
if not exist "%target_directory%" (
    echo This must be new for you...
    echo Creating your OCRAP directory here: %target_directory%
    mkdir "%target_directory%"
    echo.
)

If not exist "%icon_file%" (
    curl -o "%icon_file%" "%icon_url%" --ssl-no-revoke -s || (
            echo Error downloading icon file! 
            goto errorHalt
        )
        :: Reset errorlevel to 0
        set errorlevel=0
        echo.
        :: Check the content of the downloaded AHK file for "404: Not Found" - github download issue fix
        type "%icon_file%" | findstr /C:"404: Not Found" >nul
        if not errorlevel 1 (
            echo 404 Error occurred while downloading icon file!
        )
)

:: Check if the script file and meta file exist locally, download if not present
if not exist "%local_meta_file%" (
    goto DownloadFiles
)

if not exist "%local_script_file%" (
    goto DownloadFiles
)


:: If we got here naturally, skip download.
goto Pass1

:DownloadFiles
echo Downloading OCRAP-VA.ahk
curl -o "%local_script_file%" "%script_url%" --ssl-no-revoke -s || (
    echo Error downloading ahk file!
    goto errorHalt
)

:: Reset errorlevel to 0
set errorlevel=0

echo.

:: Check the content of the downloaded AHK file for "404: Not Found" - github download issue fix
type "%local_script_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    echo 404 Error occurred while downloading ahk script file!
    goto errorHalt
)

echo Downloading ocrap_meta.xml
curl -o "%local_meta_file%" "%meta_url%" --ssl-no-revoke -s || (
    echo Error downloading meta file!
    goto errorHalt
)

:: Reset errorlevel to 0
set errorlevel=0

echo.

:: Check the content of the downloaded Meta file for "404: Not Found" - github download issue fix
type "%local_meta_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    echo 404 Error occurred while downloading metadata file!
    echo Please report this issue to Ronald.Major@va.gov and Lewis.DeJaegher@va.gov
    pause
    echo This is just metadata to assist with updates; the script appears to be available.
    goto retryPrompt
)

:: Other situations wouldn't apply, go straight to launch
goto LaunchScript

:Pass1

:: If the file is found, check the current version
echo The launcher is comparing your local version against the latest build.
:: get remote version
echo Downloading build info...
curl -o "%temp_meta_file%" "%meta_url%" --ssl-no-revoke -s || (
    echo Error downloading Meta file!
    echo Please report this issue to Ronald.Major@va.gov
    pause
    goto retryPrompt
)

:: Reset errorlevel to 0
set errorlevel=0

:: Check the content of the downloaded Meta file for "404: Not Found" - github download issue fix
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

:: get local version
set "installedversion="
for /f "tokens=3 delims=<>" %%a in (
    'find /i "<Version>" ^< "%local_meta_file%"'
) do set "installedversion=%%a"

echo Your local OCRAP-VA.ahk version is %installedversion%
echo.

:: compare local to remote version

if "%installedversion%" EQU "%buildversion%" (
    :: if versions are the same, run the local script, you're done!
    goto LaunchScript
)

:: if versions are NOT the same, download remote before running
echo The local version is different from the build version. The latest build will now be downloaded.
curl -o "%local_script_file%" "%script_url%" --ssl-no-revoke -s || (
    echo Error downloading ahk file!
    echo Please report this issue to Ronald.Major@va.gov
    echo.
    echo Attempting to start the old version of OCRAP-VA.ahk
    goto LaunchScript
)

:: Reset errorlevel to 0
set errorlevel=0

:: Check the content of the downloaded file for "404: Not Found" - github download issue fix
type "%local_script_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    echo 404 Error occurred while downloading script file!
    goto errorHalt
)

echo.
echo Success. You are now on the latest version.

:: in any case, start the script
:LaunchScript
:: Check if AutoHotkey is installed - need to verify path
set "ahk_exe_path=C:\Program Files\AutoHotkey\v2\AutoHotkey.exe"

if not exist "%ahk_exe_path%" (
    echo.
    echo Error! - AutoHotkey v2 does not seem to be installed on this device.
    echo Autohotkey v2 is required for this script to work. This program is TRM approved.
    echo See the related Pharmacy EHRM Community of Practice Channel thread for more info.
    echo.

    :: Prompt the user for input
:retryPrompt
    set "user_choice="
    set /p "user_choice=Would you like to try to start the script anyways? Y or N: "
    :: Check the user's choice and determine the next action
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
    )
)

start "" "%local_script_file%"

:: Check if the run was successful
if %errorlevel% equ 0 (
    echo OCRAP-VA.ahk has Launched.
    echo.
) else (
    echo Launch Attempted despite some errors.
    echo Check your windows taskbar tray to confirm whether the ahk script is running.
    echo if this keeps happening, report to ronald.major@va.gov
    echo.
    pause
)
:: If no shortcuts found, prompt for creation
set "shortcut_name=OCRAP-VA"
if not exist "%startup_folder%\%shortcut_name%.lnk" (

    If not exist "%desktop_folder%\%shortcut_name%.lnk" (   
    echo It looks like you don't have any shortcuts created
    echo We recommend adding a shortcut to startup so the script starts when you log on. 
    echo Which shortcuts would you like?
    echo.
    :shortcutOptions
    echo D=Desktop
    echo S=Startup
    echo B=Both Desktop and Startup
    echo N=None
    set "user_choice_shortcut="
    set /p "user_choice_shortcut=Please select from the above options: "
:: Check the user's choice and determine the next action
    if /i "!user_choice_shortcut!"=="N" (
        echo.
        echo Exiting Launcher
        timeout /t 3
        exit /b 1
    ) else if /i "!user_choice_shortcut!"=="D" (
        If exist "%icon_file%" (
            echo Creating shortcut on the desktop...
            :: Create a shortcut on the desktop using PowerShell
            powershell -command "New-Object -ComObject WScript.Shell | ForEach-Object { $_.CreateShortcut('%desktop_folder%\%shortcut_name%.lnk') } | ForEach-Object { $_.TargetPath = '%launcher_file%'; $_.IconLocation = '%icon_file%'; $_.Save() }"
            echo Shortcut created on the desktop.
        ) else (
            echo Shortcut creation failed - no icon file
        )

    ) else if /i "!user_choice_shortcut!"=="S" (
        echo Creating shortcut in the startup folder...

        If exist "%icon_file%" (
            :: Create a shortcut using PowerShell
            powershell -command "New-Object -ComObject WScript.Shell | ForEach-Object { $_.CreateShortcut('%startup_folder%\%shortcut_name%.lnk') } | ForEach-Object { $_.TargetPath = '%launcher_file%'; $_.IconLocation = '%icon_file%'; $_.Save() }"
            echo Shortcut created in startup.
        ) else (
            echo Shortcut creation failed - no icon file
        )

    ) else if /i "!user_choice_shortcut!"=="B" (
        If exist "%icon_file%" (
            :: Create a shortcut using PowerShell
            powershell -command "New-Object -ComObject WScript.Shell | ForEach-Object { $_.CreateShortcut('%startup_folder%\%shortcut_name%.lnk') } | ForEach-Object { $_.TargetPath = '%launcher_file%'; $_.IconLocation = '%icon_file%'; $_.Save() }"
            echo Shortcut created in startup.

            :: Create a shortcut on the desktop using PowerShell
            powershell -command "New-Object -ComObject WScript.Shell | ForEach-Object { $_.CreateShortcut('%desktop_folder%\%shortcut_name%.lnk') } | ForEach-Object { $_.TargetPath = '%launcher_file%'; $_.IconLocation = '%icon_file%'; $_.Save() }"
            echo Shortcut created on the desktop.
        ) else (
            echo Shortcut creation failed - no icon file
        )
    ) else (
        echo Invalid choice.
        goto shortcutOptions
    )

)
)

if not exist %launcher_file% (
    goto downloadLauncher
)


set "launcherversion="
for /f "tokens=3 delims=<>" %%a in (
    'find /i "<Launcher>" ^< "%temp_meta_file%"'
) do set "launcherversion=%%a"

echo The current Launcher Script build is %launcherversion%

:: get local version
set "installedlauncher="
for /f "tokens=3 delims=<>" %%a in (
    'find /i "<Launcher>" ^< "%local_meta_file%"'
) do set "installedlauncher=%%a"

echo Your local Launcher Script version is %installedlauncher%
echo.

:: compare local to remote version

if "%installedlauncher%" EQU "%launcherversion%" (
    :: if versions are the same, run the local script, you're done!
    echo Launcher up to date.
    :: Copy the contents of temp_meta_file to local_meta_file, then delete the temp
    copy /y "%temp_meta_file%" "%local_meta_file%"
    del "%temp_meta_file%"
    echo all done!
    timeout /t 10
    exit /b 1
)
:downloadLauncher
curl -o "%launcher_file%" "%launcher_url%" --ssl-no-revoke -s || (
    echo Error downloading icon file!
    goto errorHalt
)

:: Reset errorlevel to 0
set errorlevel=0
echo.
:: Check the content of the downloaded file for "404: Not Found" - github download issue fix
type "%launcher_file%" | findstr /C:"404: Not Found" >nul
if not errorlevel 1 (
    for /f "usebackq delims=" %%a in ("%launcher_file%") do (
        set "first_line=%%a"
        goto :check_first_line
    )

    :check_first_line
    if "!first_line!"=="404: Not Found" (
        echo 404 Error occurred while downloading icon file!
        goto errorHalt
    )
)

echo Launcher updated
echo.
:: Copy the contents of temp_meta_file to local_meta_file, overwriting it
copy /y "%temp_meta_file%" "%local_meta_file%"
echo all done!
timeout /t 10
exit /b 1

:errorHalt
echo Please report this issue to Ronald.Major@va.gov
        pause
        exit /b 1