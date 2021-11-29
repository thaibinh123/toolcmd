CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
@echo off
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
    
		



title Cai dat Project-Visio cho may tinh!
cls
color f0
mode con: cols=60 lines=27

:MainMenu
del /s /f /q Configuration.xml
cls
set zz=
set pp=
set tt=
set mm=
set ee=
set cc=
set nn=
set vv=
set gg=
set ff=
set xx=
set yy=

:===========================================================================================================
:batdau
>> "Configuration.xml" echo ^<Configuration^>

echo. 
echo.          Chon phien ban 32bit hoac 64bit
echo.
echo.
echo.      (A): 32bit        ;         (B): 64bit
echo.
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 set xx=64
if ERRORLEVEL 1 set xx=32
>> "Configuration.xml" echo  ^<Add OfficeClientEdition="%xx%" ^>
cls
echo.
echo.        (X): Cai           ;        (Y): Khong cai
echo.
echo.
:project
echo. 1/ Cai Project?
Choice /N /C XY /M "* Nhap Lua Chon Cua Ban [X hoac Y] :
if ERRORLEVEL 2 echo. == Khong ==&goto:visio
if ERRORLEVEL 1 echo. == Co ==&set zz=ProjectPro&set vv=Project Professional

echo.
echo.   (1): phien ban 2016
echo.   (2): phien ban 2019
echo.   (3): phien ban 2021
echo.
echo.
choice /c:123 /n /m "Nhap number cua phien ban muon cai dat [1,2,3] : "
if %errorlevel% EQU 3 set pp=2021
if %errorlevel% EQU 2 set pp=2019
if %errorlevel% EQU 1 set pp=&set gg=2016_

::retail-volume
set tt=Retail



:display
>> "Configuration.xml" echo  ^<Product ID="%zz%%pp%%tt%"^>
>> "Configuration.xml" echo  ^<Language ID="en-us" /^>
>> "Configuration.xml" echo  ^</Product^>




:visio
echo.
echo. 6/ Cai Visio?
Choice /N /C XY /M "* Nhap Lua Chon Cua Ban [X hoac Y] :
if ERRORLEVEL 2 echo. == Khong ==&goto:end_all
if ERRORLEVEL 1 echo. == Co ==&set mm=VisioPro&set nn=Visio Professional

echo.
echo.   (1): phien ban 2016
echo.   (2): phien ban 2019
echo.   (3): phien ban 2021
echo.
echo.
choice /c:123 /n /m "Nhap number cua phien ban muon cai dat [1,2,3] : "
if %errorlevel% EQU 3 set cc=2021
if %errorlevel% EQU 2 set cc=2019
if %errorlevel% EQU 1 set cc=&set ff=2016_

::retail-volume
set ee=Retail


:display
>> "Configuration.xml" echo  ^<Product ID="%mm%%cc%%ee%"^>
>> "Configuration.xml" echo  ^<Language ID="en-us" /^>
>> "Configuration.xml" echo  ^</Product^>




:===========================================================================================================
:end_all
>> "Configuration.xml" echo  ^</Add^>
>> "Configuration.xml" echo  ^<Display AcceptEULA="True" /^>
>> "Configuration.xml" echo  ^<Extend CreateShortcuts="true" /^>
>> "Configuration.xml" echo  ^</Configuration^>

cls
echo.
echo.
echo.
::xet dieu kien
if [%zz%] EQU [ProjectPro] goto:chuyentiep
if [%zz%] NEQ [ProjectPro] goto:chuyentiep2

 :chuyentiep
if [%mm%] EQU [VisioPro] set chicopro=no&goto:co_pro_ne
if [%mm%] NEQ [VisioPro] set chicopro=yes&goto:co_pro_ne

:chuyentiep2
if [%mm%] NEQ [VisioPro] goto:MainMenu
if [%mm%] EQU [VisioPro] goto:co_vi_ne



::DISPLAY
:co_pro_ne
echo.      === %vv% %gg%%pp%_%tt%_%xx%bit ===
echo.
echo.
if [%chicopro%] EQU [yes] goto:endgame

:co_vi_ne
echo.      === %nn% %ff%%cc%_%ee%_%xx%bit ===
echo.
echo.






:endgame
echo.
echo.               === BAT DAU CAI DAT? ===
echo.
echo.             (Y): Yes     ;      (N): No
echo.
Choice /N /C YN /M "* Nhap Lua Chon Cua Ban [Y hoac N] :
if ERRORLEVEL 2 del /s /f /q Configuration.xml&cls&goto:MainMenu
if ERRORLEVEL 1 cls

mode con: cols=50 lines=15
echo.
echo. Dang bat dau qua trinh cai dat Project/Visio....
echo.
setup.exe /configure Configuration.xml
exit






