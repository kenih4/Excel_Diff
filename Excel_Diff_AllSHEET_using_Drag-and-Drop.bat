@echo off
setlocal

rem   .\Excel_Diff_AllSHEET_using_Drag-and-Drop_TEST.bat tmp.xlsx tmp2.xlsx

if "%~1"=="" (
    echo Please drag and drop one or more Excel files onto this batch file.
    pause
    exit /b
)


echo %~1
echo %~2
pause
python Excel_Diff_AllSHEET_using_Drag-and-Drop.py %~1 %~2
exit /b