set Dir_Old=%cd%
cd /D %~dp0

del /s /f *.ps *.dvi *.aux *.toc *.idx *.ind *.ilg *.log *.out *.brf *.blg *.bbl main.pdf

set LATEX_CMD=pdflatex
%LATEX_CMD% main
echo ----
makeindex main.idx
echo ----
%LATEX_CMD% main

setlocal enabledelayedexpansion
set count=8
:repeat
set content=X
for /F "tokens=*" %%T in ( 'findstr /C:"Rerun LaTeX" main.log' ) do set content="%%~T"
if !content! == X for /F "tokens=*" %%T in ( 'findstr /C:"Rerun to get cross-references right" main.log' ) do set content="%%~T"
if !content! == X goto :skip
set /a count-=1
if !count! EQU 0 goto :skip

echo ----
%LATEX_CMD% main
goto :repeat
:skip
endlocal
makeindex main.idx
%LATEX_CMD% main
cd /D %Dir_Old%
set Dir_Old=
