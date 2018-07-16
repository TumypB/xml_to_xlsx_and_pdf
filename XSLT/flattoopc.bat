flattoopc.exe %1 %1.xlsx
if %errorlevel% neq 0 exit /b %errorlevel%
start "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" %1.xlsx