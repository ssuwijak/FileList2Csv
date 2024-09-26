@echo off

setlocal enabledelayedexpansion

rem set "folderPath=C:\Your\Folder\Path"
set "folderPath=C:\Windows\Temp"
set "fileList=file_list.txt"

for /r "%folderPath%" %%a in (*) do (
    echo %%a >> %fileList%
)

dir *.txt
