echo off

REM Title:    Re-assembles a Microsoft Access Template file from the 'src' directory
REM Author:   Geoffrey L. Griffith
REM Date:     2025-05-16



REM Get Date and Time stamp for file name
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YY=%dt:~2,2%" & set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%" & set "Min=%dt:~10,2%" & set "Sec=%dt:~12,2%"
set "fullstamp=%YYYY%-%MM%-%DD%_%HH%.%Min%.%Sec%"
echo:
echo Timestamp: "%fullstamp%"


REM Change to the parent directory
cd ..


REM Ensure the 'bin' directory exists
if exist bin\ (
  echo:
  echo The 'bin' dir exists already 
) else (
  echo:
  echo Creating the 'bin' directory
  mkdir bin
)


REM Zipping the 'src' directory into the 'bin' folder
echo:
echo Creating Access Template in 'bin' directory...
powershell Compress-Archive -Path .\src\* -DestinationPath .\bin\TemplateBuild.zip


REM Renaming new .zip file to .accdt
echo:
echo Renaming file to .accdt...
cd bin
ren TemplateBuild.zip TemplateBuild_%fullstamp%.accdt



REM Script completed
echo:
echo ***** Template Created! *****
echo:
echo:
echo:

pause