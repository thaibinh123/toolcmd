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

====================================================================
title Ho tro cac van de ve Office cho may tinh!
mode con: cols=65 lines=25
color f0

:MainMenu
echo. 
cls
echo.
echo.                          == MENU ==
echo.      
echo.      [  1. Cai dat Word/Excel/PowerPoint...   : Nhan so 1  ] 
echo.
echo.      [  2. Cai dat Project-Visio              : Nhan so 2  ]
echo.	  
echo.      [  3. Xuat bieu tuong Office ra man hinh : Nhan so 3  ]
echo.
echo.      [  4. Go Office tan goc                  : Nhan so 4  ]
echo.	  
echo.
echo.            ----------------------------------------
echo.
echo.                  [  5 Thoat  : Nhan so 5  ]
@echo.
echo.        
echo.
@echo ===========================
Choice /N /C 12345 /M "* Nhap lua chon cua ban: 

if ERRORLEVEL 5 goto:Exit
if ERRORLEVEL 4 goto:uninstalloffice      
if ERRORLEVEL 3 goto:in_shortcut_office
if ERRORLEVEL 2 goto:installproject_visio
if ERRORLEVEL 1 goto:installoffice









:============================================================================================================
:installoffice
start office.cmd
goto:MainMenu


:============================================================================================================
:installproject_visio
start project_visio.cmd
goto:MainMenu






:============================================================================================================
:in_shortcut_office
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\W*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\E*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\P*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\V*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\A*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\O*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\T*.lnk" "%AllUsersProfile%\Desktop"
goto:MainMenu








:=========================================================================================
:uninstalloffice
cls
mode con: cols=67 lines=25
color f0
cls
echo. 
echo.                             == MENU ==
echo.      
echo.         [  1. SetupProd_OffScrub           : Nhan so 1  ]  
echo.               
echo.         [  2. o15-ctremove                 : Nhan so 2  ]
echo.
echo.         [  3. Phần mềm Your Uninstall      : Nhan so 3  ]
echo.
echo.        ---------------------------------------------------
echo.
echo.                  [  4. Quay lai   : Nhan so 4  ]
echo.
echo.
@echo ===========================
Choice /N /C 1234 /M "* Nhap lua chon cua ban : 
if ERRORLEVEL 4 goto:MainMenu
if ERRORLEVEL 3 goto:off4
if ERRORLEVEL 2 goto:off2
if ERRORLEVEL 1 goto:off1  

=========================
:off1
cls
echo.
echo.                 Tool "SetupProd_OffScrub"
echo.
echo. -La cong cu delete Office tu Microsoft, cong cu nay co the
echo. xoa "tan goc" cac phien ban Office 2010-2013-2016-2019-2021
echo  va 365, Visio va Project ban chi can an 1,2 click la đc
echo.
echo.
echo. Chu y: Tool nay se go tan goc toan bo Office bao gom ca Visio
echo. va Project, Office 365 ra khoi may. Neu chi muon xoa rieng le
echo. vd: chi muon xoa Visio? thi chon B quay lai, sau do chon so 4
echo.
echo.
echo.       [A] Remove Office         [B] Quay lai 
echo.
@echo ==============================================
Choice /N /C AB /M "* Nhap lua chon cua ban :

if ERRORLEVEL 2 goto :uninstalloffice
if ERRORLEVEL 1 start remove_office_tan_goc\SetupProd_OffScrub.exe&goto:uninstalloffice



#o15-ctremove
Là công cụ chính chủ của Microsoft dùng để gỡ bỏ hoàn toàn sạch sẽ Office ra khỏi máy tính

#Office Uninstall v1.8.2
Là công cụ gỡ tận gốc Office được tách từ tool C2R của Ratiborus (thường bị antivirus nhận nhầm là virus, vì vậy khi sd bạn hãy tạm thời off antivirus nhés)




==========================================================
:off2
cls
echo.
echo.                 Tool "o15-ctremove"
echo.
echo. -La cong cu delete Office tu Microsoft, cong cu nay co the
echo. xoa "tan goc" cac phien ban Office 2010-2013-2016-2019-2021
echo  va 365, Visio va Project ban chi can an 1,2 click la đc
echo.
echo.
echo. Chu y: Tool nay se go tan goc toan bo Office bao gom ca Visio
echo. va Project, Office 365 ra khoi may. Neu chi muon xoa rieng le
echo. vd: chi muon xoa Visio? thi chon B quay lai, sau do chon so 4
echo.
echo.
echo.       [A] Remove Office         [B] Quay lai 
echo.
@echo ==============================================
Choice /N /C AB /M "* Nhap lua chon cua ban :

if ERRORLEVEL 2 goto :uninstalloffice
if ERRORLEVEL 1 start remove_office_tan_goc\o15-ctrremove.diagcab&goto:uninstalloffice






==========================================================
:off4
cls
echo.
echo.                Phần mềm "Your Uninstaller"
echo.
echo.
echo.       [A] Download phần mềm         [B] Quay lai 
echo.
@echo ==============================================
Choice /N /C AB /M "* Nhap lua chon cua ban :
if ERRORLEVEL 2 goto :uninstalloffice
if ERRORLEVEL 1 start http://ursoftware.com/&start https://textuploader.com/1fn7k&goto:uninstalloffice



:======================================================================================================================================================
:Exit
echo. Good Bye!
timeout 3
exit
































::cach in ky tu ra file notepad thanh cong!
::Tất cả các ký hiệu Greater-Than (>), Less-Than (<), Pipe (|), Ampersand (&) và Caret (^) đều cần phải được thoát bằng dấu mũ (^) trừ khi chúng được chứa trong “dấu ngoặc kép”
::vd: echo ^)                                                                                                  ::
::vd2: echo ^> ^< ^| ^& ^^