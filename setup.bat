#! /bin/bash
set mypath=%cd%
set gitpath="%ProgramFiles%\Git\bin\"

C:\Python3.7.4\python.exe -m venv ./src/venv
%gitpath%bash.exe  -c "\"%mypath%\src\venv\Scripts\python.exe\" -m pip install --upgrade pip"
%gitpath%bash.exe  -c "\"%mypath%\src\venv\Scripts\python.exe\" -m pip install -r ./src/requirements.txt"
cd ./src/
./update_Cel.bat
Pause