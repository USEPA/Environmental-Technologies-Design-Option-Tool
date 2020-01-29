# Microsoft Developer Studio Project File - Name="stepp" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=stepp - Win32 Release
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "stepp.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "stepp.mak" CFG="stepp - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "stepp - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "stepp - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "stepp - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir ".\Release"
# PROP BASE Intermediate_Dir ".\Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ".\Release"
# PROP Intermediate_Dir ".\Release"
# PROP Target_Dir ""
F90=fl32.exe
# ADD BASE F90 /Ox /I "Release/" /c /nologo /MT
# ADD F90 /Ox /I "Release/" /c /nologo /MT
# ADD CPP /FD
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir ".\Debug"
# PROP BASE Intermediate_Dir ".\Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ".\Debug"
# PROP Intermediate_Dir ".\Debug"
# PROP Target_Dir ""
F90=fl32.exe
# ADD BASE F90 /Zi /I "Debug/" /c /nologo /MT
# ADD F90 /Zi /I "Debug/" /c /nologo /MT
# ADD CPP /FD
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386

!ENDIF 

# Begin Target

# Name "stepp - Win32 Release"
# Name "stepp - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat;for;f90"
# Begin Source File

SOURCE=.\ACCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\AIRDENS.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\AIRVISC.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\AQSCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\AQSFIT.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\AQSOL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\BINPAR.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DBDENS.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DIFGWL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DIFLHL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DIFLPOL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DIFLWC.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ERROR.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\FGRP.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\FGRPCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\GETGAM.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\H2ODENS.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\H2OST.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\H2OVISC.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\HC1CALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\HC2CALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\HCDBCONV.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\HENFIT.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\HENRY.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\INITVS.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\KOWCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\LDDBCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\LDGCCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\MOLWT.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\MWTCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\NEWTON.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ORGDENS.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\PARMS.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\PARTC.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\REGRESS.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\UNIMOD.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\VAPORP.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\VBBPCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\VBMATT.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\VBMSCH.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\VPRCALL.F90

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl;fi;fd"
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;cnt;rtf;gif;jpg;jpeg;jpe"
# End Group
# End Target
# End Project
