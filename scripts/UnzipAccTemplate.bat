echo off

REM Title:    Script to Unzip the Microsoft Access Template file
REM Author:   Geoffrey L. Griffith
REM Date:     2025-05-16


REM Change to the parent directory
cd ..


REM Delete the existing src directory and files  
echo:
echo Deleting 'src' directory...
rmdir /S /Q src


REM Create a new src directory to extract the template to 
echo:
echo Creating new 'src' directory...
mkdir src


REM Unzip the Access Template file to the new 'src' directory
REM Note: tar command supported in Windows 10 build 17063 or later  
echo:
echo Unzipping the Access template file...
cd src
tar -xf ..\accdb\Kents_FileDialog_Tools.accdt


REM Script completed
echo:
echo ***** Template Unzip Completed *****
echo:
echo:
echo:

pause

