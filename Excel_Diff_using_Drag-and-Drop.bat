@echo off
setlocal

rem   .\Excel_Diff_using_Drag-and-Drop_TEST.bat tmp.xlsx tmp2.xlsx

if "%~1"=="" (
    echo Please drag and drop one or more Excel files onto this batch file.
    pause
rem    exit /b
)


echo %~1
echo %~2
pause
python Excel_Diff_using_Drag-and-Drop_TEST.py %~1 %~2
exit /b

:loop
if "%~1"=="" goto :end
start "" "%~1"
shift
goto loop

:end
exit