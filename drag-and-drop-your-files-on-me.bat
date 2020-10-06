@ECHO off

SET toolLocation=%~dp0
ECHO Welcome to WORT
ECHO ##################
ECHO.

ECHO Do you wish to set a file format id(hit enter to use default .docx):
SET /p formatId="ID:"

:LOOP
	REM check first argument whether it is empty and quit loop in case;
	REM `%1` is the argument as is; `%~1` removes surrounding quotes;
	REM `"%~1"` ensures that the argument is always enclosed within quotes:
	IF "%~1"=="" GOTO :END
	REM the argument is passed over to the command to execute (`"%~1"`):
	ECHO.
	ECHO Processing
	ECHO %~1
	ECHO.

	REM Running wscript.exe instead of calling cscript circumnavigates the security popup for runing a VBS script
	"%SystemRoot%\System32\wscript.exe" "%toolLocation%\msword-helper.vbs"  "%~1"  "%formatId%"

	ECHO Done
	ECHO.

	REM `shift` makes the second argument (`%2`) to be the first (`%1`), the third (`%3`) to be the second (`%2`),...:
	SHIFT

	REM go back to top:
	GOTO :LOOP
:END

ECHO.
ECHO ##### All Transformations Completed #####
PAUSE & EXIT