@ECHO OFF

CD /D %~dp0

PYTHON.exe --version
ECHO IS IT OK ?
PAUSE

PIP.exe install -r requirements.txt

IF %ERRORLEVEL% NEQ 0 (
  ECHO SOMETHING WENT WRONG..
  PAUSE
)
