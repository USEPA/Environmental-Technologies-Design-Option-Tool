# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=asapptad - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to asapptad - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "asapptad - Win32 Release" && "$(CFG)" !=\
 "asapptad - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "asapptad.mak" CFG="asapptad - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "asapptad - Win32 Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "asapptad - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
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
F90=fl32.exe
MTL=mktyplib.exe

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
OUTDIR=.\Release
INTDIR=.\Release

ALL : "$(OUTDIR)\asapptad.dll"

CLEAN : 
	-@erase ".\Release\asapptad.dll"
	-@erase ".\Release\AIRDENS.OBJ"
	-@erase ".\Release\VQMLTPT1.OBJ"
	-@erase ".\Release\VQCALC.OBJ"
	-@erase ".\Release\TVOLPT2.OBJ"
	-@erase ".\Release\REMOVPT.OBJ"
	-@erase ".\Release\QH2OPT2.OBJ"
	-@erase ".\Release\QAIRPT2.OBJ"
	-@erase ".\Release\PTOTALPT.OBJ"
	-@erase ".\Release\PT1VQMIN.OBJ"
	-@erase ".\Release\PT1TVOL.OBJ"
	-@erase ".\Release\PT1LDH2O.OBJ"
	-@erase ".\Release\PT1LDAIR.OBJ"
	-@erase ".\Release\PT1HTOW.OBJ"
	-@erase ".\Release\PT1DTOW.OBJ"
	-@erase ".\Release\PT1AREA.OBJ"
	-@erase ".\Release\PPUMPPT.OBJ"
	-@erase ".\Release\PDROP.OBJ"
	-@erase ".\Release\PBLOWPT.OBJ"
	-@erase ".\Release\OPTMAL.OBJ"
	-@erase ".\Release\ONDKLAPT.OBJ"
	-@erase ".\Release\ONDAKLPT.OBJ"
	-@erase ".\Release\ONDAKGPT.OBJ"
	-@erase ".\Release\LDH2OPT2.OBJ"
	-@erase ".\Release\LDAIRPT2.OBJ"
	-@erase ".\Release\KLACOR.OBJ"
	-@erase ".\Release\H2OVISC.OBJ"
	-@erase ".\Release\H2OST.OBJ"
	-@erase ".\Release\H2ODENS.OBJ"
	-@erase ".\Release\GETSAF.OBJ"
	-@erase ".\Release\GETNTUPT.OBJ"
	-@erase ".\Release\GETMULT.OBJ"
	-@erase ".\Release\GETHTUPT.OBJ"
	-@erase ".\Release\GETHIVQ.OBJ"
	-@erase ".\Release\GETCSPT.OBJ"
	-@erase ".\Release\FINDKLA.OBJ"
	-@erase ".\Release\EFFLPT2.OBJ"
	-@erase ".\Release\DIFLPOL.OBJ"
	-@erase ".\Release\DIFLHL.OBJ"
	-@erase ".\Release\DIFGWL.OBJ"
	-@erase ".\Release\DIFFL.OBJ"
	-@erase ".\Release\AWCALC.OBJ"
	-@erase ".\Release\AREAPT2.OBJ"
	-@erase ".\Release\AIRVISC.OBJ"
	-@erase ".\Release\AIRFLO.OBJ"
	-@erase ".\Release\asapptad.lib"
	-@erase ".\Release\asapptad.exp"

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
BSC32_FLAGS=/nologo /o"$(OUTDIR)/asapptad.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)/asapptad.pdb" /machine:I386 /out:"$(OUTDIR)/asapptad.dll"\
 /implib:"$(OUTDIR)/asapptad.lib" 
LINK32_OBJS= \
	"$(INTDIR)/AIRDENS.OBJ" \
	"$(INTDIR)/VQMLTPT1.OBJ" \
	"$(INTDIR)/VQCALC.OBJ" \
	"$(INTDIR)/TVOLPT2.OBJ" \
	"$(INTDIR)/REMOVPT.OBJ" \
	"$(INTDIR)/QH2OPT2.OBJ" \
	"$(INTDIR)/QAIRPT2.OBJ" \
	"$(INTDIR)/PTOTALPT.OBJ" \
	"$(INTDIR)/PT1VQMIN.OBJ" \
	"$(INTDIR)/PT1TVOL.OBJ" \
	"$(INTDIR)/PT1LDH2O.OBJ" \
	"$(INTDIR)/PT1LDAIR.OBJ" \
	"$(INTDIR)/PT1HTOW.OBJ" \
	"$(INTDIR)/PT1DTOW.OBJ" \
	"$(INTDIR)/PT1AREA.OBJ" \
	"$(INTDIR)/PPUMPPT.OBJ" \
	"$(INTDIR)/PDROP.OBJ" \
	"$(INTDIR)/PBLOWPT.OBJ" \
	"$(INTDIR)/OPTMAL.OBJ" \
	"$(INTDIR)/ONDKLAPT.OBJ" \
	"$(INTDIR)/ONDAKLPT.OBJ" \
	"$(INTDIR)/ONDAKGPT.OBJ" \
	"$(INTDIR)/LDH2OPT2.OBJ" \
	"$(INTDIR)/LDAIRPT2.OBJ" \
	"$(INTDIR)/KLACOR.OBJ" \
	"$(INTDIR)/H2OVISC.OBJ" \
	"$(INTDIR)/H2OST.OBJ" \
	"$(INTDIR)/H2ODENS.OBJ" \
	"$(INTDIR)/GETSAF.OBJ" \
	"$(INTDIR)/GETNTUPT.OBJ" \
	"$(INTDIR)/GETMULT.OBJ" \
	"$(INTDIR)/GETHTUPT.OBJ" \
	"$(INTDIR)/GETHIVQ.OBJ" \
	"$(INTDIR)/GETCSPT.OBJ" \
	"$(INTDIR)/FINDKLA.OBJ" \
	"$(INTDIR)/EFFLPT2.OBJ" \
	"$(INTDIR)/DIFLPOL.OBJ" \
	"$(INTDIR)/DIFLHL.OBJ" \
	"$(INTDIR)/DIFGWL.OBJ" \
	"$(INTDIR)/DIFFL.OBJ" \
	"$(INTDIR)/AWCALC.OBJ" \
	"$(INTDIR)/AREAPT2.OBJ" \
	"$(INTDIR)/AIRVISC.OBJ" \
	"$(INTDIR)/AIRFLO.OBJ"

"$(OUTDIR)\asapptad.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

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
OUTDIR=.\Debug
INTDIR=.\Debug

ALL : "$(OUTDIR)\asapptad.dll"

CLEAN : 
	-@erase ".\Debug\asapptad.dll"
	-@erase ".\Debug\AIRDENS.OBJ"
	-@erase ".\Debug\VQMLTPT1.OBJ"
	-@erase ".\Debug\VQCALC.OBJ"
	-@erase ".\Debug\TVOLPT2.OBJ"
	-@erase ".\Debug\REMOVPT.OBJ"
	-@erase ".\Debug\QH2OPT2.OBJ"
	-@erase ".\Debug\QAIRPT2.OBJ"
	-@erase ".\Debug\PTOTALPT.OBJ"
	-@erase ".\Debug\PT1VQMIN.OBJ"
	-@erase ".\Debug\PT1TVOL.OBJ"
	-@erase ".\Debug\PT1LDH2O.OBJ"
	-@erase ".\Debug\PT1LDAIR.OBJ"
	-@erase ".\Debug\PT1HTOW.OBJ"
	-@erase ".\Debug\PT1DTOW.OBJ"
	-@erase ".\Debug\PT1AREA.OBJ"
	-@erase ".\Debug\PPUMPPT.OBJ"
	-@erase ".\Debug\PDROP.OBJ"
	-@erase ".\Debug\PBLOWPT.OBJ"
	-@erase ".\Debug\OPTMAL.OBJ"
	-@erase ".\Debug\ONDKLAPT.OBJ"
	-@erase ".\Debug\ONDAKLPT.OBJ"
	-@erase ".\Debug\ONDAKGPT.OBJ"
	-@erase ".\Debug\LDH2OPT2.OBJ"
	-@erase ".\Debug\LDAIRPT2.OBJ"
	-@erase ".\Debug\KLACOR.OBJ"
	-@erase ".\Debug\H2OVISC.OBJ"
	-@erase ".\Debug\H2OST.OBJ"
	-@erase ".\Debug\H2ODENS.OBJ"
	-@erase ".\Debug\GETSAF.OBJ"
	-@erase ".\Debug\GETNTUPT.OBJ"
	-@erase ".\Debug\GETMULT.OBJ"
	-@erase ".\Debug\GETHTUPT.OBJ"
	-@erase ".\Debug\GETHIVQ.OBJ"
	-@erase ".\Debug\GETCSPT.OBJ"
	-@erase ".\Debug\FINDKLA.OBJ"
	-@erase ".\Debug\EFFLPT2.OBJ"
	-@erase ".\Debug\DIFLPOL.OBJ"
	-@erase ".\Debug\DIFLHL.OBJ"
	-@erase ".\Debug\DIFGWL.OBJ"
	-@erase ".\Debug\DIFFL.OBJ"
	-@erase ".\Debug\AWCALC.OBJ"
	-@erase ".\Debug\AREAPT2.OBJ"
	-@erase ".\Debug\AIRVISC.OBJ"
	-@erase ".\Debug\AIRFLO.OBJ"
	-@erase ".\Debug\asapptad.ilk"
	-@erase ".\Debug\asapptad.lib"
	-@erase ".\Debug\asapptad.exp"
	-@erase ".\Debug\asapptad.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "Debug/" /c /nologo /MT
# ADD F90 /Zi /I "Debug/" /c /nologo /MT
F90_PROJ=/Zi /I "Debug/" /c /nologo /MT /Fo"Debug/" /Fd"Debug/asapptad.pdb" 
F90_OBJS=.\Debug/
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/asapptad.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)/asapptad.pdb" /debug /machine:I386\
 /out:"$(OUTDIR)/asapptad.dll" /implib:"$(OUTDIR)/asapptad.lib" 
LINK32_OBJS= \
	"$(INTDIR)/AIRDENS.OBJ" \
	"$(INTDIR)/VQMLTPT1.OBJ" \
	"$(INTDIR)/VQCALC.OBJ" \
	"$(INTDIR)/TVOLPT2.OBJ" \
	"$(INTDIR)/REMOVPT.OBJ" \
	"$(INTDIR)/QH2OPT2.OBJ" \
	"$(INTDIR)/QAIRPT2.OBJ" \
	"$(INTDIR)/PTOTALPT.OBJ" \
	"$(INTDIR)/PT1VQMIN.OBJ" \
	"$(INTDIR)/PT1TVOL.OBJ" \
	"$(INTDIR)/PT1LDH2O.OBJ" \
	"$(INTDIR)/PT1LDAIR.OBJ" \
	"$(INTDIR)/PT1HTOW.OBJ" \
	"$(INTDIR)/PT1DTOW.OBJ" \
	"$(INTDIR)/PT1AREA.OBJ" \
	"$(INTDIR)/PPUMPPT.OBJ" \
	"$(INTDIR)/PDROP.OBJ" \
	"$(INTDIR)/PBLOWPT.OBJ" \
	"$(INTDIR)/OPTMAL.OBJ" \
	"$(INTDIR)/ONDKLAPT.OBJ" \
	"$(INTDIR)/ONDAKLPT.OBJ" \
	"$(INTDIR)/ONDAKGPT.OBJ" \
	"$(INTDIR)/LDH2OPT2.OBJ" \
	"$(INTDIR)/LDAIRPT2.OBJ" \
	"$(INTDIR)/KLACOR.OBJ" \
	"$(INTDIR)/H2OVISC.OBJ" \
	"$(INTDIR)/H2OST.OBJ" \
	"$(INTDIR)/H2ODENS.OBJ" \
	"$(INTDIR)/GETSAF.OBJ" \
	"$(INTDIR)/GETNTUPT.OBJ" \
	"$(INTDIR)/GETMULT.OBJ" \
	"$(INTDIR)/GETHTUPT.OBJ" \
	"$(INTDIR)/GETHIVQ.OBJ" \
	"$(INTDIR)/GETCSPT.OBJ" \
	"$(INTDIR)/FINDKLA.OBJ" \
	"$(INTDIR)/EFFLPT2.OBJ" \
	"$(INTDIR)/DIFLPOL.OBJ" \
	"$(INTDIR)/DIFLHL.OBJ" \
	"$(INTDIR)/DIFGWL.OBJ" \
	"$(INTDIR)/DIFFL.OBJ" \
	"$(INTDIR)/AWCALC.OBJ" \
	"$(INTDIR)/AREAPT2.OBJ" \
	"$(INTDIR)/AIRVISC.OBJ" \
	"$(INTDIR)/AIRFLO.OBJ"

"$(OUTDIR)\asapptad.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
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

# Name "asapptad - Win32 Release"
# Name "asapptad - Win32 Debug"

!IF  "$(CFG)" == "asapptad - Win32 Release"

!ELSEIF  "$(CFG)" == "asapptad - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\AIRDENS.F90

"$(INTDIR)\AIRDENS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VQMLTPT1.F90

"$(INTDIR)\VQMLTPT1.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VQCALC.F90

"$(INTDIR)\VQCALC.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\TVOLPT2.F90

"$(INTDIR)\TVOLPT2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\REMOVPT.F90

"$(INTDIR)\REMOVPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\QH2OPT2.F90

"$(INTDIR)\QH2OPT2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\QAIRPT2.F90

"$(INTDIR)\QAIRPT2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PTOTALPT.F90

"$(INTDIR)\PTOTALPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PT1VQMIN.F90

"$(INTDIR)\PT1VQMIN.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PT1TVOL.F90

"$(INTDIR)\PT1TVOL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PT1LDH2O.F90

"$(INTDIR)\PT1LDH2O.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PT1LDAIR.F90

"$(INTDIR)\PT1LDAIR.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PT1HTOW.F90

"$(INTDIR)\PT1HTOW.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PT1DTOW.F90

"$(INTDIR)\PT1DTOW.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PT1AREA.F90

"$(INTDIR)\PT1AREA.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PPUMPPT.F90

"$(INTDIR)\PPUMPPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PDROP.F90

"$(INTDIR)\PDROP.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PBLOWPT.F90

"$(INTDIR)\PBLOWPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\OPTMAL.F90

"$(INTDIR)\OPTMAL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ONDKLAPT.F90

"$(INTDIR)\ONDKLAPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ONDAKLPT.F90

"$(INTDIR)\ONDAKLPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ONDAKGPT.F90

"$(INTDIR)\ONDAKGPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDH2OPT2.F90

"$(INTDIR)\LDH2OPT2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDAIRPT2.F90

"$(INTDIR)\LDAIRPT2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\KLACOR.F90

"$(INTDIR)\KLACOR.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\H2OVISC.F90

"$(INTDIR)\H2OVISC.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\H2OST.F90

"$(INTDIR)\H2OST.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\H2ODENS.F90

"$(INTDIR)\H2ODENS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETSAF.F90

"$(INTDIR)\GETSAF.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETNTUPT.F90

"$(INTDIR)\GETNTUPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETMULT.F90

"$(INTDIR)\GETMULT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETHTUPT.F90

"$(INTDIR)\GETHTUPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETHIVQ.F90

"$(INTDIR)\GETHIVQ.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GETCSPT.F90

"$(INTDIR)\GETCSPT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\FINDKLA.F90

"$(INTDIR)\FINDKLA.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\EFFLPT2.F90

"$(INTDIR)\EFFLPT2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\DIFLPOL.F90

"$(INTDIR)\DIFLPOL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\DIFLHL.F90

"$(INTDIR)\DIFLHL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\DIFGWL.F90

"$(INTDIR)\DIFGWL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\DIFFL.F90

"$(INTDIR)\DIFFL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AWCALC.F90

"$(INTDIR)\AWCALC.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AREAPT2.F90

"$(INTDIR)\AREAPT2.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AIRVISC.F90

"$(INTDIR)\AIRVISC.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AIRFLO.F90

"$(INTDIR)\AIRFLO.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
# End Target
# End Project
################################################################################
