
set BaseDir="C:\\VisualK\\Proyectos\\Factura Electronica VK 2"
set BinDir="C:\\VisualK\\Proyectos\\Factura Electronica VK 2\\Bin\\Debug"
set BinFile="C:\\VisualK\\Proyectos\\Factura Electronica VK 2\\Bin\\Debug\\Factura Electronica.exe"
set Version=1.003.09
call "%VS110COMNTOOLS%vsvars32.bat"

msbuild "Factura Electronica VK.sln" /t:Clean,Build /p:Configuration=Debug;Platform=x86
set BUILD_STATUS=%ERRORLEVEL%
if %BUILD_STATUS%==0 GOTO Reactor
pause
EXIT

:Reactor
"C:\\Program Files (x86)\\Eziriz\\.NET Reactor\\dotNET_Reactor.exe" -project "C:\\VisualK\\Proyectos\\Factura Electronica VK 2\\Bin\\Debug\\Factura Electronica.nrproj" -targetfile "C:\\VisualK\\Proyectos\\Factura Electronica VK 2\\Bin\\Debug\\Factura Electronica.exe"
set REACTOR_STATUS=%ERRORLEVEL%
if %REACTOR_STATUS%==0 GOTO INNO
pause
EXIT

:INNO
"C:\Program Files (x86)\Inno Setup 5\iscc.exe" ""C:\\VisualK\\Proyectos\\Factura Electronica VK 2\\Factura Electronica VK.iss"
set INNO_STATUS=%ERRORLEVEL%
if %INNO_STATUS%==0 GOTO ARD
pause


