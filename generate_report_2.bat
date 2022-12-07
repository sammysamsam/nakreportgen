set mypath=%cd%
set gitpath="%ProgramFiles%\Git\bin\"


rd /s /q "%mypath%\DONE"
mkdir "%mypath%\DONE"

cd src/process_files_2
run.bat