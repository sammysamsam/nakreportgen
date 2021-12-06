#! /bin/bash
set mypath=%cd%
set gitpath="%ProgramFiles%\Git\bin\"

C:\Python3.7.4\python.exe -m venv ./report_generator/venv
%gitpath%bash.exe  -c "\"%mypath%\report_generator\venv\Scripts\python.exe\" -m pip install --upgrade pip"
%gitpath%bash.exe  -c "\"%mypath%\report_generator\venv\Scripts\python.exe\" -m pip install -r ./report_generator/requirements.txt"
Pause