@echo off
setlocal
title Audit Analytics Workbench - Install
cd /d "%~dp0"
echo.
echo   Audit Analytics Workbench - one-time installation
echo   -------------------------------------------------
echo   This copies two libraries (DuckDB and openpyxl) from the "wheels"
echo   folder into Python on this computer. No internet connection is used.
echo.

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto nopython

%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>nul || goto oldpython
%PY% -c "import tkinter" >nul 2>nul || goto notk

%PY% -m pip install --user --no-index --find-links "%~dp0wheels" duckdb openpyxl
if errorlevel 1 goto failed
%PY% -c "import duckdb, openpyxl" >nul 2>nul || goto failed

echo.
echo   Installation complete.
echo   Now double-click "Run Audit Workbench".
echo.
pause
exit /b 0

:nopython
echo   Python was not found on this computer.
echo   Please ask IT to install Python 3.11 or newer (64-bit) from python.org,
echo   with the default options. Then double-click Install again.
goto end

:oldpython
echo   The Python on this computer is older than 3.11.
echo   Please ask IT to install Python 3.11 or newer (64-bit) from python.org.
goto end

:notk
echo   Python is installed without its window toolkit (tcl/tk).
echo   Please ask IT to re-run the Python installer, choose "Modify" and tick
echo   "tcl/tk and IDLE". Then double-click Install again.
goto end

:failed
echo.
echo   Installation did not complete. Please take a photo of this window
echo   and send it to the person who supports the workbench.

:end
echo.
pause
exit /b 1
