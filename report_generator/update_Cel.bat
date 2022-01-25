#! /bin/bash

set mypath=%cd%
del"%mypath%\venv\Lib\site-packages\pdf2docx\table\Cell.py
COPY Cell.py "%mypath%\venv\Lib\site-packages\pdf2docx\table\Cell.py
pause 100 