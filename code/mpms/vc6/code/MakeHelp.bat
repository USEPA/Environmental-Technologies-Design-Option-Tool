@echo off
REM -- First make map file from Microsoft Visual C++ generated resource.h
echo // MAKEHELP.BAT generated Help Map file.  Used by UPPMEM.HPJ. >"hlp\Uppmem.hm"
echo. >>"hlp\Uppmem.hm"
echo // Commands (ID_* and IDM_*) >>"hlp\Uppmem.hm"
makehm ID_,HID_,0x10000 IDM_,HIDM_,0x10000 resource.h >>"hlp\Uppmem.hm"
echo. >>"hlp\Uppmem.hm"
echo // Prompts (IDP_*) >>"hlp\Uppmem.hm"
makehm IDP_,HIDP_,0x30000 resource.h >>"hlp\Uppmem.hm"
echo. >>"hlp\Uppmem.hm"
echo // Resources (IDR_*) >>"hlp\Uppmem.hm"
makehm IDR_,HIDR_,0x20000 resource.h >>"hlp\Uppmem.hm"
echo. >>"hlp\Uppmem.hm"
echo // Dialogs (IDD_*) >>"hlp\Uppmem.hm"
makehm IDD_,HIDD_,0x20000 resource.h >>"hlp\Uppmem.hm"
echo. >>"hlp\Uppmem.hm"
echo // Frame Controls (IDW_*) >>"hlp\Uppmem.hm"
makehm IDW_,HIDW_,0x50000 resource.h >>"hlp\Uppmem.hm"
REM -- Make help for Project UPPMEM


echo Building Win32 Help files
start /wait hcw /C /E /M "hlp\Uppmem.hpj"
if errorlevel 1 goto :Error
if not exist "hlp\Uppmem.hlp" goto :Error
if not exist "hlp\Uppmem.cnt" goto :Error
echo.
if exist Debug\nul copy "hlp\Uppmem.hlp" Debug
if exist Debug\nul copy "hlp\Uppmem.cnt" Debug
if exist Release\nul copy "hlp\Uppmem.hlp" Release
if exist Release\nul copy "hlp\Uppmem.cnt" Release
echo.
goto :done

:Error
echo hlp\Uppmem.hpj(1) : error: Problem encountered creating help file

:done
echo.
