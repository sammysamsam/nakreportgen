#! /bin/bash
set mypath=%cd%
set gitpath="%ProgramFiles%\Git\bin\"

rd /s /q "%mypath%\..\..\DONE"
mkdir "%mypath%\..\..\DONE"

%gitpath%bash.exe  -c "\"%mypath%\..\venv\Scripts\python.exe\" main.py -debug=1"

Pause