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

		

:Version: 2.0
:Developer: Thanos
:OS support [32+64bit]: Windows 7/8/8.1 (chỉ cài được Office 2010, 2013, 2016 Volume), Windows 10 (cài được mọi bản), Windows 11 (cài được mọi bản)

:========================================================================================================
:MainMenu
title Cai dat Word,Excel,Powerpoint... cho may tinh!
color f0
mode con: cols=57


:startok
del /s /f /q Configuration.xml
cls
set aa=
set bb=
set xx=
set yy=
set off365=
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
>> "Configuration.xml" echo ^<Configuration^>
echo.
echo. ===  Lua chon phien ban Office ban muon cai dat ===
echo.
echo.
ECHO              1. Office 2010 Pro Plus
ECHO              -----------------------
ECHO              2. Office 2013 Pro Plus
ECHO              -----------------------
ECHO              3. Office 2016 Pro Plus
ECHO              -----------------------
ECHO              4. Office 2019 Pro Plus 
ECHO              -----------------------
ECHO              5. Office 2021 Pro Plus 
ECHO              -----------------------
ECHO              6. Office 365 Pro Plus 
echo.
echo.
echo. -----------------------
choice /c:123456 /n /m "Chon phien ban muon cai dat [1,2,3,4,5,6] : "
if %errorlevel% EQU 1 set aa=2010&set yy=Office Professional Plus 2010&goto:1
if %errorlevel% EQU 2 set aa=2013&set yy=Office Professional Plus 2013&goto:1
if %errorlevel% EQU 3 set aa=ProPlus&set yy=Office Professional Plus 2016&goto:1
if %errorlevel% EQU 4 set aa=ProPlus2019&set yy=Office Professional Plus 2019&goto:1
if %errorlevel% EQU 5 set aa=ProPlus2021&set yy=Office Professional Plus 2021&goto:1
if %errorlevel% EQU 6 set aa=O365ProPlus&set yy=Office 365&set off365=ok&goto:1
if %errorlevel% NEQ 1 goto:startok
if %errorlevel% NEQ 2 goto:startok
if %errorlevel% NEQ 3 goto:startok
if %errorlevel% NEQ 4 goto:startok
if %errorlevel% NEQ 5 goto:startok
if %errorlevel% NEQ 6 goto:startok


:1
echo.
echo.      (A): 32bit     ;      (B): 64bit
echo.
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 set xx=64
if ERRORLEVEL 1 set xx=32
>> "Configuration.xml" echo  ^<Add OfficeClientEdition="%xx%" ^>


::retail-volume
if [%off365%] EQU [ok] set bb=Retail&goto:tiepdi
echo.
echo.    (R): Retail     ;      (V): Volume
echo.
Choice /N /C RV /M "* Nhap Lua Chon Cua Ban [R hoac V] :
if ERRORLEVEL 2 set bb=Volume
if ERRORLEVEL 1 set bb=Retail

:tiepdi
::display
if [%aa%] EQU [2010] goto:download
if [%aa%] EQU [2013] goto:download
if [%aa%] EQU [ProPlus] set aa=2016&goto:2016nha
goto:display

:2016nha
if [%bb%] EQU [Volume] goto:download
if [%bb%] EQU [Retail] cls

:display
>> "Configuration.xml" echo  ^<Product ID="%aa%%bb%"^>
cls







::Option_App
:part1
>> "Configuration.xml" echo  ^<Language ID="en-us" /^>
echo.
echo.
echo.    ___Lua chon phan mem ma ban muon cai!____
echo.
echo.
echo.       A: Cai      ;      B: Khong cai
echo.
echo. 1/ Word?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem1
if ERRORLEVEL 1 echo. == Co ==&set a=Word&goto:part2
:lem1
>> "Configuration.xml" echo  ^<ExcludeApp ID="Word" /^> 

:part2
echo.   
echo. 2/ Excel?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem2
if ERRORLEVEL 1 echo. == Co ==&set b= + Excel&goto:part3
:lem2
>> "Configuration.xml" echo  ^<ExcludeApp ID="Excel" /^> 

:part3
echo.
echo. 3/ PowerPoint?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem3
if ERRORLEVEL 1 echo. == Co ==&set c= + PowerPoint&goto:part4
:lem3
>> "Configuration.xml" echo  ^<ExcludeApp ID="PowerPoint" /^> 

:part4
echo.
echo. 4/ Access?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem4
if ERRORLEVEL 1 echo. == Co ==&set d= + Access&goto:part5
:lem4
>> "Configuration.xml" echo  ^<ExcludeApp ID="Access" /^> 


:part5
echo.
echo. 5/ Publisher?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem5
if ERRORLEVEL 1 echo. == Co ==&set e= + Publisher&goto:part6
:lem5
>> "Configuration.xml" echo  ^<ExcludeApp ID="Publisher" /^> 


:part6
echo.
echo. 6/ Outlook?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem6
if ERRORLEVEL 1 echo. == Co ==&set f= + Outlook&goto:part7
:lem6
>> "Configuration.xml" echo  ^<ExcludeApp ID="Outlook" /^> 


:part7
echo.
echo. 7/ OneNote?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem7
if ERRORLEVEL 1 echo. == Co ==&set g= + OneNote&goto:part8
:lem7
>> "Configuration.xml" echo  ^<ExcludeApp ID="OneNote" /^> 



:part8
echo.
echo. 8/ OneDrive?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem8
if ERRORLEVEL 1 echo. == Co ==&set h= + OneDrive&goto:part9
:lem8
>> "Configuration.xml" echo  ^<ExcludeApp ID="OneDrive" /^> 


:part9
if [%off365%] EQU [ok] goto:tieptuc
if [%off365%] NEQ [ok] goto:endoffice
:tieptuc
echo.
echo. 9/ Microsoft Teams?
Choice /N /C AB /M "* Nhap Lua Chon Cua Ban [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem9
if ERRORLEVEL 1 echo. == Co ==&set k= + Teams&goto:endoffice
:lem9
>> "Configuration.xml" echo  ^<ExcludeApp ID="Teams" /^> 


:endoffice
>> "Configuration.xml" echo  ^<ExcludeApp ID="Groove" /^> 
>> "Configuration.xml" echo  ^<ExcludeApp ID="Lync" /^> 
>> "Configuration.xml" echo  ^</Product^>

:===========================================================================================================
:part12
echo.
echo.
echo.      ==================================
echo.               Project - Visio
echo.      ==================================
echo.
echo.       X: Cai        ;     Y: Khong cai
echo.
echo. 10/ Project Pro?
Choice /N /C XY /M "* Nhap Lua Chon Cua Ban [X hoac Y] :
if ERRORLEVEL 2 echo. == Khong ==&goto:part13
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




:part13
echo.
echo. 11/ Visio Pro?
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
echo.     === %yy%_%bb%_%xx%bit ===
echo.  (%a%%b%%c%%d%%e%%f%%g%%h%%i%%j%%k%)
echo.
echo.
if [%zz%] NEQ [ProjectPro] goto:buocnhay
echo.      === %vv% %gg%%pp%_%tt%_%xx%bit ===
echo.
echo.
:buocnhay
if [%mm%] NEQ [VisioPro] goto:buocnhay2
echo.      === %nn% %ff%%cc%_%ee%_%xx%bit ===
echo.
echo.
echo.
echo.
:buocnhay2
echo.               === BAT DAU CAI DAT? ===
echo.
echo.             (Y): Yes     ;      (N): No
echo.
Choice /N /C YN /M "* Nhap Lua Chon Cua Ban [Y hoac N] :
if ERRORLEVEL 2 del /s /f /q Configuration.xml&cls&goto:startok
if ERRORLEVEL 1 cls

mode con: cols=50 lines=15
echo.
echo. Dang bat dau qua trinh cai dat Office....
echo.
setup.exe /configure Configuration.xml
exit











:download
mode con: cols=62 lines=20
if [%aa%] EQU [2010] goto:2010 
if [%aa%] EQU [2013] goto:2013
if [%aa%] EQU [2016] goto:2016ne


:2010
if [%xx%] NEQ [32] goto:64bitne
if [%xx%] EQU [32] cls
if [%bb%] EQU [Retail] start https://gdtxbadinh-my.sharepoint.com/:u:/g/personal/billgates_gdtxbadinh_onmicrosoft_com/Ee6hZYK5Fp1JnjEemmMz0jgBGYMznpy1tCWUyyZ1eZgzfw.
if [%bb%] EQU [Volume] start https://kichhoat-my.sharepoint.com/:u:/g/personal/365_kichhoat_onmicrosoft_com/EVPKuwf26udMqyQZ0iIyD2gBjFwkzaU_L8ROkJQrxcjQYA?e=uFOpnq
goto:tieptheo
:64bitne
if [%bb%] EQU [Retail] start https://kichhoat-my.sharepoint.com/:u:/g/personal/365_kichhoat_onmicrosoft_com/EccDWGErlchEketRpKIh6swBBbKOhhEm0C_lwlu8lsWaRg?e=A0e6XZ
if [%bb%] EQU [Volume] start https://gdtxbadinh-my.sharepoint.com/:u:/g/personal/billgates_gdtxbadinh_onmicrosoft_com/EbSt7_AEE5tHreJBQID7UxcBQII1hCh3Urb0lxBN1bPiXw.
goto:tieptheo


:2013
if [%xx%] NEQ [32] goto:64bitnha
if [%xx%] EQU [32] cls
if [%bb%] EQU [Retail] start https://kichhoat-my.sharepoint.com/:u:/g/personal/365_kichhoat_onmicrosoft_com/ERKjtbW25K9Mjkx9ns_WHygBnJFceOpmToTLzZ-tq0IX-w?e=sceOeh
if [%bb%] EQU [Volume] start https://kichhoat-my.sharepoint.com/:u:/g/personal/365_kichhoat_onmicrosoft_com/EW35RV16BxFKi6LxGQcTV7MBYChxKpm3Pvu4yG9o3fGu6A?e=xaaX7f
goto:tieptheo
:64bitnha
if [%bb%] EQU [Retail] start https://kichhoat-my.sharepoint.com/:u:/g/personal/365_kichhoat_onmicrosoft_com/EX9W13khU4VLrb_Z7Z29sAABOmeTIadjqTcde-QdD4Gfrg?e=gDX83g
if [%bb%] EQU [Volume] start https://kichhoat-my.sharepoint.com/:u:/g/personal/365_kichhoat_onmicrosoft_com/EQdEpCrGX3ZPpltlJ_gY6PABh8OJixKy-hSwlhDk9yj_VQ?e=e5hZa0
goto:tieptheo


:2016ne
if [%xx%] EQU [32] start https://gdtxbadinh-my.sharepoint.com/:u:/g/personal/billgates_gdtxbadinh_onmicrosoft_com/EUlgt5vqD_lDnHf8b7idhHQBki1vJL0vNutembjdDkd4ig.
if [%xx%] EQU [64] start https://gdtxbadinh-my.sharepoint.com/:u:/g/personal/billgates_gdtxbadinh_onmicrosoft_com/EZ2rx98vpgZBqiKk5vzr-u4BhYwggC7y__yUabeQpWqINA.
goto:tieptheo


:tieptheo
cls
echo.
echo.
echo.
echo.        === %yy%_%bb%_%xx%bit ===
echo.
echo.
echo.
echo. 1/ Cac ban download file theo link
echo. 2/ Sau khi tai xong, cac ban click 2 lan vao file vua tai ve
echo.    roi click vao file "setup" 2 lan roi bat dau cai nha
echo.
echo.
echo.
echo. Nhan phim bat ky de quay lai MENU...
pause >nul
start Office.cmd
exit