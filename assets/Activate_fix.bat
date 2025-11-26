@echo off
echo Office 2021 Activator - Fixed Version
echo =====================================
echo.

REM Check and navigate to Office16 directory (64-bit)
if exist "C:\Program Files\Microsoft Office\root\Office16\ospp.vbs" (
    cd /d "C:\Program Files\Microsoft Office\root\Office16"
    echo Found Office 64-bit installation
    goto :activate
)

REM Check and navigate to Office16 directory (32-bit)
if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\ospp.vbs" (
    cd /d "C:\Program Files (x86)\Microsoft Office\root\Office16"
    echo Found Office 32-bit installation
    goto :activate
)

echo ERROR: Office installation not found!
echo Please check if Office 2021 is installed correctly.
pause
exit /b 1

:activate
echo.
echo Installing licenses...
for /f %%x in ('dir /b ..\Licenses16\ProPlus2021VL_KMS*.xrm-ms 2^>nul') do (
    echo Installing: %%x
    cscript ospp.vbs /inslic:"..\Licenses16\%%x"
)

echo.
echo Installing product key...
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

echo.
echo Setting KMS server...
cscript ospp.vbs /sethst:kms.msgang.com

echo.
echo Activating Office...
cscript ospp.vbs /act

echo.
echo =====================================
echo Activation process completed!
echo =====================================
pause

