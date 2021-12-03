#! /bin/bash
set mypath=%cd%
set gitpath="%ProgramFiles%\Git\bin\"

C:\Python3.7.4\python.exe -m venv ./venv
%gitpath%bash.exe  -c "\"%mypath%\venv\Scripts\python.exe\" -m pip install --upgrade pip"
%gitpath%bash.exe  -c "\"%mypath%\venv\Scripts\python.exe\" -m pip install -r requirements.txt"
