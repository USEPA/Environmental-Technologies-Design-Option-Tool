# Microsoft Developer Studio Project File - Name="stepp_test1" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=stepp_test1 - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "stepp_test1.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "stepp_test1.mak" CFG="stepp_test1 - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "stepp_test1 - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "stepp_test1 - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
F90=df.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "stepp_test1 - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Target_Dir ""
# ADD BASE F90 /compile_only /include:"Release/" /dll /libs:dll /nologo /warn:nofileopt
# ADD F90 /compile_only /include:"Release/" /dll /libs:dll /nologo /warn:nofileopt
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /machine:I386

!ELSEIF  "$(CFG)" == "stepp_test1 - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Target_Dir ""
# ADD BASE F90 /check:bounds /compile_only /debug:full /include:"Debug/" /dll /libs:dll /nologo /traceback /warn:argument_checking /warn:nofileopt
# ADD F90 /check:bounds /compile_only /debug:full /include:"Debug/" /dll /libs:dll /nologo /traceback /warn:argument_checking /warn:nofileopt
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /YX /FD /GZ   /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /YX /FD /GZ   /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /incremental:no /debug /machine:I386 /pdbtype:sept

!ENDIF 

# Begin Target

# Name "stepp_test1 - Win32 Release"
# Name "stepp_test1 - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat;f90;for;f;fpp"
# Begin Source File

SOURCE=.\ACCALL.F90
# End Source File
# Begin Source File

SOURCE=.\AIRDENS.F90
# End Source File
# Begin Source File

SOURCE=.\AIRVISC.F90
# End Source File
# Begin Source File

SOURCE=.\AQSCALL.F90
# End Source File
# Begin Source File

SOURCE=.\AQSFIT.F90
# End Source File
# Begin Source File

SOURCE=.\AQSOL.F90
# End Source File
# Begin Source File

SOURCE=.\BINPAR.F90
# End Source File
# Begin Source File

SOURCE=.\DBDENS.F90
# End Source File
# Begin Source File

SOURCE=.\DIFGWL.F90
# End Source File
# Begin Source File

SOURCE=.\DIFLHL.F90
# End Source File
# Begin Source File

SOURCE=.\DIFLPOL.F90
# End Source File
# Begin Source File

SOURCE=.\DIFLWC.F90
# End Source File
# Begin Source File

SOURCE=.\ERROR.F90
# End Source File
# Begin Source File

SOURCE=.\FGRP.F90
# End Source File
# Begin Source File

SOURCE=.\FGRPCALL.F90
# End Source File
# Begin Source File

SOURCE=.\GETGAM.F90
# End Source File
# Begin Source File

SOURCE=.\H2ODENS.F90
# End Source File
# Begin Source File

SOURCE=.\H2OST.F90
# End Source File
# Begin Source File

SOURCE=.\H2OVISC.F90
# End Source File
# Begin Source File

SOURCE=.\HC1CALL.F90
# End Source File
# Begin Source File

SOURCE=.\HC2CALL.F90
# End Source File
# Begin Source File

SOURCE=.\HCDBCONV.F90
# End Source File
# Begin Source File

SOURCE=.\HENFIT.F90
# End Source File
# Begin Source File

SOURCE=.\HENRY.F90
# End Source File
# Begin Source File

SOURCE=.\INITVS.F90
# End Source File
# Begin Source File

SOURCE=.\KOWCALL.F90
# End Source File
# Begin Source File

SOURCE=.\LDDBCALL.F90
# End Source File
# Begin Source File

SOURCE=.\LDGCCALL.F90
# End Source File
# Begin Source File

SOURCE=.\MOLWT.F90
# End Source File
# Begin Source File

SOURCE=.\MWTCALL.F90
# End Source File
# Begin Source File

SOURCE=.\NEWTON.F90
# End Source File
# Begin Source File

SOURCE=.\ORGDENS.F90
# End Source File
# Begin Source File

SOURCE=.\PARMS.F90
# End Source File
# Begin Source File

SOURCE=.\PARTC.F90
# End Source File
# Begin Source File

SOURCE=.\REGRESS.F90
# End Source File
# Begin Source File

SOURCE=.\UNIMOD.F90
# End Source File
# Begin Source File

SOURCE=.\VAPORP.F90
# End Source File
# Begin Source File

SOURCE=.\VBBPCALL.F90
# End Source File
# Begin Source File

SOURCE=.\VBMATT.F90
# End Source File
# Begin Source File

SOURCE=.\VBMSCH.F90
# End Source File
# Begin Source File

SOURCE=.\VPRCALL.F90
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl;fi;fd"
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# End Group
# End Target
# End Project
