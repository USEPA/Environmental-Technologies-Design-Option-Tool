# Microsoft Developer Studio Project File - Name="asapptad" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=asapptad - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "asapptad.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "asapptad.mak" CFG="asapptad - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "asapptad - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "asapptad - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
F90=df.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "asapptad - Win32 Release"

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

!ELSEIF  "$(CFG)" == "asapptad - Win32 Debug"

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

# Name "asapptad - Win32 Release"
# Name "asapptad - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat;f90;for;f;fpp"
# Begin Source File

SOURCE=.\AIRDENS.F90
# End Source File
# Begin Source File

SOURCE=.\AIRFLO.F90
# End Source File
# Begin Source File

SOURCE=.\AIRVISC.F90
# End Source File
# Begin Source File

SOURCE=.\AREAPT2.F90
# End Source File
# Begin Source File

SOURCE=.\AWCALC.F90
# End Source File
# Begin Source File

SOURCE=.\DIFFL.F90
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

SOURCE=.\EFFLPT2.F90
# End Source File
# Begin Source File

SOURCE=.\FINDKLA.F90
# End Source File
# Begin Source File

SOURCE=.\GETCSPT.F90
# End Source File
# Begin Source File

SOURCE=.\GETHIVQ.F90
# End Source File
# Begin Source File

SOURCE=.\GETHTUPT.F90
# End Source File
# Begin Source File

SOURCE=.\GETMULT.F90
# End Source File
# Begin Source File

SOURCE=.\GETNTUPT.F90
# End Source File
# Begin Source File

SOURCE=.\GETSAF.F90
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

SOURCE=.\KLACOR.F90
# End Source File
# Begin Source File

SOURCE=.\LDAIRPT2.F90
# End Source File
# Begin Source File

SOURCE=.\LDH2OPT2.F90
# End Source File
# Begin Source File

SOURCE=.\ONDAKGPT.F90
# End Source File
# Begin Source File

SOURCE=.\ONDAKLPT.F90
# End Source File
# Begin Source File

SOURCE=.\ONDKLAPT.F90
# End Source File
# Begin Source File

SOURCE=.\OPTMAL.F90
# End Source File
# Begin Source File

SOURCE=.\PBLOWPT.F90
# End Source File
# Begin Source File

SOURCE=.\PDROP.F90
# End Source File
# Begin Source File

SOURCE=.\PPUMPPT.F90
# End Source File
# Begin Source File

SOURCE=.\PT1AREA.F90
# End Source File
# Begin Source File

SOURCE=.\PT1DTOW.F90
# End Source File
# Begin Source File

SOURCE=.\PT1HTOW.F90
# End Source File
# Begin Source File

SOURCE=.\PT1LDAIR.F90
# End Source File
# Begin Source File

SOURCE=.\PT1LDH2O.F90
# End Source File
# Begin Source File

SOURCE=.\PT1TVOL.F90
# End Source File
# Begin Source File

SOURCE=.\PT1VQMIN.F90
# End Source File
# Begin Source File

SOURCE=.\PTOTALPT.F90
# End Source File
# Begin Source File

SOURCE=.\QAIRPT2.F90
# End Source File
# Begin Source File

SOURCE=.\QH2OPT2.F90
# End Source File
# Begin Source File

SOURCE=.\REMOVPT.F90
# End Source File
# Begin Source File

SOURCE=.\TVOLPT2.F90
# End Source File
# Begin Source File

SOURCE=.\VQCALC.F90
# End Source File
# Begin Source File

SOURCE=.\VQMLTPT1.F90
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
