@echo off
REM run_validate.bat - clean installer and runner for the validator

cd /d "%~dp0"

echo Checking for Python (searching for 'py' or 'python')...
where py >nul 2>&1
if errorlevel 1 (
	where python >nul 2>&1
	if errorlevel 1 (
		echo Python not found in PATH.
		echo Please install Python and select 'Add Python to PATH', then try again.
		pause
		exit /b 1
	) else (
		set "PY_CMD=python"
	)
) else (
	set "PY_CMD=py"
)

echo Checking dependencies in requirements.txt...
if not exist requirements.txt (
	echo requirements.txt not found. Skipping package installation.
) else (
	for /f "usebackq tokens=* delims=" %%A in ("requirements.txt") do (
		call :process_req "%%A"
	)
)

echo Running validator...
rem If no arguments were provided, interactively ask for the spreadsheet path first, then ask for start row.
if "%*"=="" (
	rem Ask for spreadsheet path (user may press Enter to let the Python script use its default)
	set /p "INPUT_PATH=Digite o caminho para a planilha .xlsx ou pressione Enter para usar o padrão: "
	rem After the path is provided (or default is chosen), ask for the starting row
	set /p "ROW=Digite o número da linha para iniciar ou pressione Enter para processar toda a planilha: "

	rem strip surrounding quotes if user pasted the path/row with quotes
	if defined INPUT_PATH set "INPUT_PATH=%INPUT_PATH:"=%"
	if defined ROW set "ROW=%ROW:"=%"
	if "%INPUT_PATH%"=="" (
		rem No explicit path provided: call Python without positional path so it can prompt/use default
		if defined ROW (
			echo Running validator for default sheet starting at row %ROW%...
			%PY_CMD% validate_pdf_standard.py --start-row %ROW%
		) else (
			echo Running validator for default sheet...
			%PY_CMD% validate_pdf_standard.py
		)
	) else (
		rem Path provided by user; quote it to allow spaces
		if defined ROW (
			echo Running validator for sheet "%INPUT_PATH%" starting at row %ROW%...
			%PY_CMD% validate_pdf_standard.py "%INPUT_PATH%" --start-row %ROW%
		) else (
			echo Running validator for sheet "%INPUT_PATH%" - full...
			%PY_CMD% validate_pdf_standard.py "%INPUT_PATH%"
		)
	)
) else (
	rem Arguments provided: pass them through to the Python script unchanged
	%PY_CMD% validate_pdf_standard.py %*
)

echo.
echo Press any key to close...
pause >nul

goto :eof

:process_req
setlocal
set "line=%~1"
for /f "tokens=* delims= " %%T in ("%line%") do set "line=%%T"
if "%line%"=="" (
	endlocal & goto :eof
)
if "%line:~0,1%"=="#" (
	endlocal & goto :eof
)
echo Checking package: %line%
%PY_CMD% -m pip show %line% >nul 2>&1
if errorlevel 1 (
	echo Installing %line% ...
	%PY_CMD% -m pip install --user %line%
) else (
	echo Package %line% already installed. Skipping.
)
endlocal & goto :eof

