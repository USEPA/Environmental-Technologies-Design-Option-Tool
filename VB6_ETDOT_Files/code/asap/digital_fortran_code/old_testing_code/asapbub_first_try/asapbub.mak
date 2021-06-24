# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=asapbub - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to asapbub - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "asapbub - Win32 Release" && "$(CFG)" !=\
 "asapbub - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "asapbub.mak" CFG="asapbub - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "asapbub - Win32 Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "asapbub - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 
################################################################################
# Begin Project
RSC=rc.exe
MTL=mktyplib.exe
F90=fl32.exe

!IF  "$(CFG)" == "asapbub - Win32 Release"

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
OUTDIR=.\Release
INTDIR=.\Release

ALL : "$(OUTDIR)\asapbub.dll"

CLEAN : 
	-@erase ".\Release\asapbub.dll"
	-@erase ".\Release\VQMINBUB.OBJ"
	-@erase ".\Release\VQBUB.OBJ"
	-@erase ".\Release\VOLBUB.OBJ"
	-@erase ".\Release\TRUEKLA.OBJ"
	-@erase ".\Release\TAUSVOLS.OBJ"
	-@erase ".\Release\REMOVBUB.OBJ"
	-@erase ".\Release\PCALCBUB.OBJ"
	-@erase ".\Release\KLABUB.OBJ"
	-@erase ".\Release\KLA20A.OBJ"
	-@erase ".\Release\GETSOTR.OBJ"
	-@erase ".\Release\GETSOTE.OBJ"
	-@erase ".\Release\GETPHIB.OBJ"
	-@erase ".\Release\GETCSTAR.OBJ"
	-@erase ".\Release\EFFLBUB.OBJ"
	-@erase ".\Release\DIFO2.OBJ"
	-@erase ".\Release\AIRFLO.OBJ"
	-@erase ".\Release\asapbub.lib"
	-@erase ".\Release\asapbub.exp"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Ox /I "Release/" /c /nologo /MT
# ADD F90 /Ox /I "Release/" /c /nologo /MT
F90_PROJ=/Ox /I "Release/" /c /nologo /MT /Fo"Release/" 
F90_OBJS=.\Release/
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /win32
MTL_PROJ=/nologo /D "NDEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/asapbub.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)/asapbub.pdb" /machine:I386 /out:"$(OUTDIR)/asapbub.dll"\
 /implib:"$(OUTDIR)/asapbub.lib" 
LINK32_OBJS= \
	"$(INTDIR)/VQMINBUB.OBJ" \
	"$(INTDIR)/VQBUB.OBJ" \
	"$(INTDIR)/VOLBUB.OBJ" \
	"$(INTDIR)/TRUEKLA.OBJ" \
	"$(INTDIR)/TAUSVOLS.OBJ" \
	"$(INTDIR)/REMOVBUB.OBJ" \
	"$(INTDIR)/PCALCBUB.OBJ" \
	"$(INTDIR)/KLABUB.OBJ" \
	"$(INTDIR)/KLA20A.OBJ" \
	"$(INTDIR)/GETSOTR.OBJ" \
	"$(INTDIR)/GETSOTE.OBJ" \
	"$(INTDIR)/GETPHIB.OBJ" \
	"$(INTDIR)/GETCSTAR.OBJ" \
	"$(INTDIR)/EFFLBUB.OBJ" \
	"$(INTDIR)/DIFO2.OBJ" \
	"$(INTDIR)/AIRFLO.OBJ"

"$(OUTDIR)\asapbub.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "asapbub - Win32 Debug"

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
OUTDIR=.\Debug
INTDIR=.\Debug

ALL : "$(OUTDIR)\asapbub.dll"

CLEAN : 
	-@erase ".\Debug\asapbub.dll"
	-@erase ".\Debug\VQMINBUB.OBJ"
	-@erase ".\Debug\VQBUB.OBJ"
	-@erase ".\Debug\VOLBUB.OBJ"
	-@erase ".\Debug\TRUEKLA.OBJ"
	-@erase ".\Debug\TAUSVOLS.OBJ"
	-@erase ".\Debug\REMOVBUB.OBJ"
	-@erase ".\Debug\PCALCBUB.OBJ"
	-@erase ".\Debug\KLABUB.OBJ"
	-@erase ".\Debug\KLA20A.OBJ"
	-@erase ".\Debug\GETSOTR.OBJ"
	-@erase ".\Debug\GETSOTE.OBJ"
	-@erase ".\Debug\GETPHIB.OBJ"
	-@erase ".\Debug\GETCSTAR.OBJ"
	-@erase ".\Debug\EFFLBUB.OBJ"
	-@erase ".\Debug\DIFO2.OBJ"
	-@erase ".\Debug\AIRFLO.OBJ"
	-@erase ".\Debug\asapbub.ilk"
	-@erase ".\Debug\asapbub.lib"
	-@erase ".\Debug\asapbub.exp"
	-@erase ".\Debug\asapbub.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "Debug/" /c /nologo /MT
# ADD F90 /Zi /I "Debug/" /c /nologo /MT
F90_PROJ=/Zi /I "Debug/" /c /nologo /MT /Fo"Debug/" /Fd"Debug/asapbub.pdb" 
F90_OBJS=.\Debug/
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/asapbub.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)/asapbub.pdb" /debug /machine:I386 /out:"$(OUTDIR)/asapbub.dll"\
 /implib:"$(OUTDIR)/asapbub.lib" 
LINK32_OBJS= \
	"$(INTDIR)/VQMINBUB.OBJ" \
	"$(INTDIR)/VQBUB.OBJ" \
	"$(INTDIR)/VOLBUB.OBJ" \
	"$(INTDIR)/TRUEKLA.OBJ" \
	"$(INTDIR)/TAUSVOLS.OBJ" \
	"$(INTDIR)/REMOVBUB.OBJ" \
	"$(INTDIR)/PCALCBUB.OBJ" \
	"$(INTDIR)/KLABUB.OBJ" \
	"$(INTDIR)/KLA20A.OBJ" \
	"$(INTDIR)/GETSOTR.OBJ" \
	"$(INTDIR)/GETSOTE.OBJ" \
	"$(INTDIR)/GETPHIB.OBJ" \
	"$(INTDIR)/GETCSTAR.OBJ" \
	"$(INTDIR)/EFFLBUB.OBJ" \
	"$(INTDIR)/DIFO2.OBJ" \
	"$(INTDIR)/AIRFLO.OBJ"

"$(OUTDIR)\asapbub.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ENDIF 

.for{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f90{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

################################################################################
# Begin Target

# Name "asapbub - Win32 Release"
# Name "asapbub - Win32 Debug"

!IF  "$(CFG)" == "asapbub - Win32 Release"

!ELSEIF  "$(CFG)" == "asapbub - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\VQMINBUB.F90

"$(INTDIR)\VQMINBUB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VQBUB.F90

"$(INTDIR)\VQBUB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VOLBUB.F90

"$(INTDIR)\VOLBUB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\TRUEKLA.F90

"$(INTDIR)\TRUEKLA.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\TAUSVOLS.F90

"$(INTDIR)\TAUSVOLS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\REMOVBUB.F90

"$(INTDIR)\REMOVBUB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PCALCBUB.F90

"$(INTDIR)\PCALCBUB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\KLABUB.F90

"$(INTDIR)\KLABUB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\KLA20A.F90

"$(INTDIR)\KLA20A.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETSOTR.F90

"$(INTDIR)\GETSOTR.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETSOTE.F90

"$(INTDIR)\GETSOTE.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETPHIB.F90

"$(INTDIR)\GETPHIB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETCSTAR.F90

"$(INTDIR)\GETCSTAR.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\EFFLBUB.F90

"$(INTDIR)\EFFLBUB.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\DIFO2.F90

"$(INTDIR)\DIFO2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AIRFLO.F90

"$(INTDIR)\AIRFLO.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
# End Target
# End Project
################################################################################
