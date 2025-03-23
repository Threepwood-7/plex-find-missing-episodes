@ECHO OFF

CHCP 65001

CD /D %~dp0

DEL /F /Q zz_*.log *.xlsx 2>nul

ECHO:
ECHO:
ECHO PLEASE WAIT..
ECHO:
ECHO:

PYTHON.exe pl_report_missing_episodes_claude.py 2> zz_pl_report_missing_episodes_claude_stderr.log

IF %ERRORLEVEL% NEQ 0 (
    ECHO:
    ECHO:
    ECHO ERROR, PLEASE CHECK THE ERROR LOG [ zz_pl_report_missing_episodes_claude_stderr.log ]
    ECHO:
    ECHO:
    PAUSE
    EXIT 0
)

ECHO:
ECHO:
ECHO DONE
ECHO:
ECHO:

TIMEOUT 5
