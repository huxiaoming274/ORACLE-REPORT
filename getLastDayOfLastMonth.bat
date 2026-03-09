@echo off
setlocal

:: Test the function
call :getLastDayOfLastMonth YYYYMMDD
call :getLastDayOfLastMonth YYYY-MM-DD
call :getLastDayOfLastMonth MM/DD/YYYY



:: Function to get the last day of the last month
:getLastDayOfLastMonth
set "format=%1"
set "year="
set "month="
set "day="

:: Get the current date components
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set datetime=%%i
set "year=%datetime:~0,4%"
set "month=%datetime:~4,2%"
set "day=%datetime:~6,2%"

:: Adjust the month and year to get the last month
if "%month%"=="01" (
    set /a year-=1
    set month=12
) else (
    set /a month-=1
    if %month% lss 10 set month=0%month%
)

:: Get the last day of the last month
for /f "tokens=2 delims==" %%i in ('"wmic path win32_localtime where (year=%year% and month=%month%) get dayofweek /value"') do set dayofweek=%%i
set /a lastDay=31
if %month%==04 set lastDay=30
if %month%==06 set lastDay=30
if %month%==09 set lastDay=30
if %month%==11 set lastDay=30
if %month%==02 (
    set /a lastDay=28
    set /a leapYear=year %% 4
    if %leapYear%==0 (
        set /a leapYear=year %% 100
        if %leapYear% neq 0 set lastDay=29
        set /a leapYear=year %% 400
        if %leapYear%==0 set lastDay=29
    )
)

:: Format the date based on the input format
set "result="
if "%format%"=="YYYYMMDD" set result=%year%%month%%lastDay%
if "%format%"=="YYYY-MM-DD" set result=%year%-%month%-%lastDay%
if "%format%"=="MM/DD/YYYY" set result=%month%/%lastDay%/%year%

:: Return the result
echo %result%
rem exit /b
pause

