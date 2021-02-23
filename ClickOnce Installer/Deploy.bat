@echo off

:: use parameters: %0
::  %~1 Name (and folder)
::  %~2 Version 
::  %~3 Icon file name (optional)

:: =====================
::		NOTES
::
::	--- Referenced Assemblies ---
::	For this to run correctly without errors, you may need to copy the applicable referenced
::	assemblies from the development machine, where Visual Studio was originally used to compile 
::	the launcher application. They would typically be in the following location:
::		C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework
::	See the following links for more details:
::		http://stackoverflow.com/questions/17220615/where-can-i-download-the-net-4-5-multitargeting-pack-for-my-build-server
::		http://stackoverflow.com/questions/10006012/how-to-get-rid-of-msbuild-warning-msb3644
::
::	--- Certificate Installation ---
::	The cert.pfx file does not need to be installed under the current user in order to sign the 
::	installation, but the user will encounter a security warning if they attempt to install
::	the application without the certificate installed as a CA.
::
:: =====================

:setvars

:: Change to parent of script directory
pushd %~dp0..

:: Get path for msbuild to rebuild launcher application using new name.
for /D %%D in (%SYSTEMROOT%\Microsoft.NET\Framework\v4*) do set msbuild.exe=%%D\MSBuild.exe
::echo %msbuild.exe%

:: Replace periods with underscores for publish path  (1.2.3.4 becomes 1_2_3_4)
set verpath=%~2
set verpath=%verpath:.=_%
::echo %verpath%

:: Use default icon if none specified
set appicon=%~3
if "%appicon%"=="" goto :copydefaulticon
set iconpath=..\..\%~1\%~2\
copy "%~1\%~2\%appicon%" _Tools\Launcher\
goto :startbuild

:copydefaulticon
copy /y _Tools\Launcher\app.ico "%~1\%~2\"
set appicon=iblp-db.ico

:startbuild

:: Build the launcher application using the application name and version
call %msbuild.exe% /target:publish _Tools\Launcher\Launcher.sln /property:configuration=RELEASE;ApplicationVersion=%~2;AssemblyName="%~1";ApplicationIcon="%appicon%" /target:Build;Clean

:: Check for errors
IF %errorlevel% NEQ 0 GOTO :error

:: Remove custom icon from launcher project
if "%appicon%" NEQ "app.ico" del _Tools\Launcher\%appicon%

:: Copy manifest files 
copy /y "_Tools\Launcher\bin\Release\app.publish\Application Files\%~1_%verpath%\%~1.exe" "%~1\%~2\"
copy /y "_Tools\Launcher\bin\Release\app.publish\Application Files\%~1_%verpath%\%~1.exe.manifest" "%~1\%~2\"
copy /y "_Tools\Launcher\bin\Release\app.publish\Application Files\%~1_%verpath%\%~1.exe.config" "%~1\%~2\"
copy /y "_Tools\Launcher\bin\Release\app.publish\Application Files\%~1_%verpath%\app.ico" "%~1\%~2\"

:: Create deployment manifest from template, if it does not exist.
echo n | copy /-y _Tools\Template.application "%~1.application"

:: goto :done 
:updatemanifests

:: Add any new files and update version
echo.
echo Processing files...
echo. 

:: Add project files to application manifest
_Tools\mage -Update "%~1\%~2\%~1.exe.manifest" -FromDirectory "%~1\%~2" -IconFile "%appicon%"

:: Sign with stored cert 
_Tools\mage -Sign "%~1\%~2\%~1.exe.manifest" -CertFile _Tools\cert.pfx

:: Update and sign deployment
_Tools\mage -Update "%~1.application" -AppManifest "%~1\%~2\%~1.exe.manifest" -Publisher IBLP -Name "%~1" -Version %~2 -MinVersion %~2 -ProviderURL "file://\\server\share\Apps\Deploy\%~1.application"
_Tools\mage -Sign "%~1.application" -CertFile _Tools\cert.pfx

echo.
echo Process Complete.
echo.
goto :done

:error

echo.
echo Houston, we have a problem.

:done

pause