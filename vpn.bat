for /F "tokens=2 delims=[." %%f in ('ver') do SET  KELLYWINVERDETECTED=%%f
set KELLYWINVERDETECTED=%KELLYWINVERDETECTED:~-1%
IF .%PROCESSOR_ARCHITECTURE% ==.IA64 GOTO INSTALL64
IF .%PROCESSOR_ARCHITECTURE% ==.AMD64 GOTO INSTALL64
IF .%PROCESSOR_ARCHITEW6432% ==.AMD64 GOTO INSTALL64
IF .%PROCESSOR_ARCHITEW6432% ==.IA64 GOTO INSTALL64

:INSTALL32
IF %KELLYWINVERDETECTED%==6 GOTO 32BIT-VISTA-7-2008
IF %KELLYWINVERDETECTED%==5 GOTO 32BIT-2000-XP-2003
GOTO OSNOTFOUND

:INSTALL64
IF %KELLYWINVERDETECTED%==6 GOTO 64BIT-VISTA-7-2008
IF %KELLYWINVERDETECTED%==5 GOTO 64BIT-XP-2003
GOTO OSNOTFOUND

:32BIT-VISTA-7-2008
REM WMIC OS GET CAPTION | findstr /i /C:"7 Professional"
REM IF %errorlevel%==0 GOTO 32BIT-7-UNSUPPORTED
WMIC OS GET CAPTION | findstr /i /C:"7 Home"
IF %errorlevel%==0 GOTO 32BIT-7-UNSUPPORTED
WMIC OS GET CAPTION | findstr /i /C:"7"
IF %errorlevel%==0 GOTO :32BIT-7
WMIC OS GET CAPTION | findstr /i /C:"Vista"
IF %errorlevel%==0 GOTO 32BIT-VISTA
GOTO OSNOTFOUND

:32BIT-VISTA
rasphone -r "KELLY VPN"
rasphone -r "KAIG VPN"
IF NOT EXIST "%appdata%\KELLY VPN Setup\" MD "%appdata%\Kelly VPN Setup\"
copy /Y "%~dp0/VPNSetup_X86_(Win7+Vista).exe" "%appdata%\KELLY VPN Setup\"
echo start /D "%appdata%\KELLY VPN Setup" "VPNSetup" VPNSetup_X86_(Win7+Vista).exe /q:a /c:"cmstp.exe VPNSetup.inf /s /su" > "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
echo del %%0 >> "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
msiexec /i "%~dp0RSA EAP Client 7.0.msi" /quiet
shutdown -r -t 0
GOTO END

:32BIT-7
rasphone -r "KELLY VPN"
rasphone -r "KAIG VPN"
IF NOT EXIST "%appdata%\KELLY VPN Setup\" MD "%appdata%\Kelly VPN Setup\"
copy /Y "%~dp0/VPNSetup_X86_(Win7+Vista).exe" "%appdata%\KELLY VPN Setup\"
echo start /D "%appdata%\KELLY VPN Setup" "VPNSetup" VPNSetup_X86_(Win7+Vista).exe /q:a /c:"cmstp.exe VPNSetup.inf /s /su" > "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
echo del %%0 >> "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
msiexec /i "%~dp0RSA EAP Client 7.1(x86).msi" /quiet
shutdown -r -t 0
GOTO END

:32BIT-2000-XP-2003
rasphone -r "KELLY VPN"
rasphone -r "KAIG VPN"
IF NOT EXIST "%appdata%\KELLY VPN Setup\" MD "%appdata%\Kelly VPN Setup\"
copy /Y "%~dp0/VPNSetup.exe" "%appdata%\KELLY VPN Setup\"
echo start /D "%appdata%\KELLY VPN Setup" "VPNSetup" VPNSetup.exe /q:a /c:"cmstp.exe VPNSetup.inf /s /su" > "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
echo del %%0 >> "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
msiexec /qn /I "%~dp0RSA Authentication Agent for Windows.rsa"
GOTO END

:64BIT-VISTA-7-2008
WMIC OS GET CAPTION | Find /i /C:"home"
IF "%ERRORLEVEL%"=="0" GOTO 64BIT-VISTA-7-HOME
rasphone -r "Kelly VPN"
rasphone -r "KAIG VPN"
IF NOT EXIST "%appdata%\KELLY VPN Setup\" MD "%appdata%\Kelly VPN Setup\"
copy /Y "%~dp0/VPNSetup_X64_(Win7+Vista).exe" "%appdata%\KELLY VPN Setup\"
echo start /D "%appdata%\KELLY VPN Setup" "VPNSetup" VPNSetup_X64_(Win7+Vista).exe /q:a /c:"cmstp.exe VPNSetup.inf /s /su" > "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
echo del %%0 >> "%userprofile%\Start Menu\Programs\Startup\RSA.bat"
msiexec /i "%~dp0RSA EAP Client 7.1.msi" /quiet ENABLESHORTCUT=0
shutdown -r -t 0
GOTO END

:64BIT-XP-2003
prompt.vbs "64-bit Windows XP/2003 are not supported. Please contact the systems department for a list of supported operating systems."
GOTO END

:64BIT-VISTA-7-HOME
prompt.vbs "64-bit Windows Vista/7 Home Edition are not supported. Please contact the systems department for a list of supported operating systems."
GOTO END

:32BIT-7-UNSUPPORTED
prompt.vbs "32-bit Windows 7 Home/Professional Editions are not supported. Please contact the systems department for a list of supported operating systems."
GOTO END

:OSNOTFOUND
prompt.vbs "Please contact the systems department. Error: OS Not Found. Installation Halted."
GOTO END

:END