set mypath=%cd%
set gitpath="%ProgramFiles%\Git\bin\"


rd /s /q "%mypath%\DONE"
mkdir "%mypath%\DONE"


rd /s /q "%mypath%\tmp"
mkdir "%mypath%\tmp"

%gitpath%bash.exe  -c "\"%mypath%\report_generator\venv\Scripts\python.exe\" ./report_generator/main.py -mode=2"

rd /s /q "%mypath%\tmp"
Pause