# Microsoft Developer Studio Project File - Name="stepconv_test1" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=stepconv_test1 - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "stepconv_test1.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "stepconv_test1.mak" CFG="stepconv_test1 - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "stepconv_test1 - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "stepconv_test1 - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
F90=df.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "stepconv_test1 - Win32 Release"

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

!ELSEIF  "$(CFG)" == "stepconv_test1 - Win32 Debug"

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

# Name "stepconv_test1 - Win32 Release"
# Name "stepconv_test1 - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat;f90;for;f;fpp"
# Begin Source File

SOURCE=.\ACCONV.F90
# End Source File
# Begin Source File

SOURCE=.\ACENSI.F90
# End Source File
# Begin Source File

SOURCE=.\ADENENSI.F90
# End Source File
# Begin Source File

SOURCE=.\ADENSCNV.F90
# End Source File
# Begin Source File

SOURCE=.\AQSCONV.F90
# End Source File
# Begin Source File

SOURCE=.\AQSENSI.F90
# End Source File
# Begin Source File

SOURCE=.\AVISCCNV.F90
# End Source File
# Begin Source File

SOURCE=.\AVISENSI.F90
# End Source File
# Begin Source File

SOURCE=.\GDIFENSI.F90
# End Source File
# Begin Source File

SOURCE=.\GDIFFCNV.F90
# End Source File
# Begin Source File

SOURCE=.\H2OSTCNV.F90
# End Source File
# Begin Source File

SOURCE=.\HCCONV.F90
# End Source File
# Begin Source File

SOURCE=.\HCENSI.F90
# End Source File
# Begin Source File

SOURCE=.\KOWCONV.F90
# End Source File
# Begin Source File

SOURCE=.\KOWENSI.F90
# End Source File
# Begin Source File

SOURCE=.\LDENENSI.F90
# End Source File
# Begin Source File

SOURCE=.\LDENSCNV.F90
# End Source File
# Begin Source File

SOURCE=.\LDIFENSI.F90
# End Source File
# Begin Source File

SOURCE=.\LDIFFCNV.F90
# End Source File
# Begin Source File

SOURCE=.\MVBPENSI.F90
# End Source File
# Begin Source File

SOURCE=.\MVNBPCNV.F90
# End Source File
# Begin Source File

SOURCE=.\MVOPTCNV.F90
# End Source File
# Begin Source File

SOURCE=.\MVOTENSI.F90
# End Source File
# Begin Source File

SOURCE=.\MWCONV.F90
# End Source File
# Begin Source File

SOURCE=.\MWENSI.F90
# End Source File
# Begin Source File

SOURCE=.\NBPCONV.F90
# End Source File
# Begin Source File

SOURCE=.\NBPENSI.F90
# End Source File
# Begin Source File

SOURCE=.\PRESENSI.F90
# End Source File
# Begin Source File

SOURCE=.\PRESSCNV.F90
# End Source File
# Begin Source File

SOURCE=.\RICONV.F90
# End Source File
# Begin Source File

SOURCE=.\RIENSI.F90
# End Source File
# Begin Source File

SOURCE=.\TEMPCNV.F90
# End Source File
# Begin Source File

SOURCE=.\TEMPENSI.F90
# End Source File
# Begin Source File

SOURCE=.\VPCONV.F90
# End Source File
# Begin Source File

SOURCE=.\VPENSI.F90
# End Source File
# Begin Source File

SOURCE=.\WDENENSI.F90
# End Source File
# Begin Source File

SOURCE=.\WDENSCNV.F90
# End Source File
# Begin Source File

SOURCE=.\WSTENSI.F90
# End Source File
# Begin Source File

SOURCE=.\WVISCCNV.F90
# End Source File
# Begin Source File

SOURCE=.\WVISENSI.F90
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
