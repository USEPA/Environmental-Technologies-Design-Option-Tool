# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=stepconv - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to stepconv - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "stepconv - Win32 Release" && "$(CFG)" !=\
 "stepconv - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "stepconv.mak" CFG="stepconv - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "stepconv - Win32 Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "stepconv - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
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

!IF  "$(CFG)" == "stepconv - Win32 Release"

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

ALL : "$(OUTDIR)\stepconv.dll"

CLEAN : 
	-@erase ".\Release\stepconv.dll"
	-@erase ".\Release\WVISENSI.OBJ"
	-@erase ".\Release\WVISCCNV.OBJ"
	-@erase ".\Release\WSTENSI.OBJ"
	-@erase ".\Release\WDENSCNV.OBJ"
	-@erase ".\Release\WDENENSI.OBJ"
	-@erase ".\Release\VPENSI.OBJ"
	-@erase ".\Release\VPCONV.OBJ"
	-@erase ".\Release\TEMPENSI.OBJ"
	-@erase ".\Release\TEMPCNV.OBJ"
	-@erase ".\Release\RIENSI.OBJ"
	-@erase ".\Release\RICONV.OBJ"
	-@erase ".\Release\PRESSCNV.OBJ"
	-@erase ".\Release\PRESENSI.OBJ"
	-@erase ".\Release\NBPENSI.OBJ"
	-@erase ".\Release\NBPCONV.OBJ"
	-@erase ".\Release\MWENSI.OBJ"
	-@erase ".\Release\MWCONV.OBJ"
	-@erase ".\Release\MVOTENSI.OBJ"
	-@erase ".\Release\MVOPTCNV.OBJ"
	-@erase ".\Release\MVNBPCNV.OBJ"
	-@erase ".\Release\MVBPENSI.OBJ"
	-@erase ".\Release\LDIFFCNV.OBJ"
	-@erase ".\Release\LDIFENSI.OBJ"
	-@erase ".\Release\LDENSCNV.OBJ"
	-@erase ".\Release\LDENENSI.OBJ"
	-@erase ".\Release\KOWENSI.OBJ"
	-@erase ".\Release\KOWCONV.OBJ"
	-@erase ".\Release\HCENSI.OBJ"
	-@erase ".\Release\HCCONV.OBJ"
	-@erase ".\Release\H2OSTCNV.OBJ"
	-@erase ".\Release\GDIFFCNV.OBJ"
	-@erase ".\Release\GDIFENSI.OBJ"
	-@erase ".\Release\AVISENSI.OBJ"
	-@erase ".\Release\AVISCCNV.OBJ"
	-@erase ".\Release\AQSENSI.OBJ"
	-@erase ".\Release\AQSCONV.OBJ"
	-@erase ".\Release\ADENSCNV.OBJ"
	-@erase ".\Release\ADENENSI.OBJ"
	-@erase ".\Release\ACENSI.OBJ"
	-@erase ".\Release\ACCONV.OBJ"
	-@erase ".\Release\stepconv.lib"
	-@erase ".\Release\stepconv.exp"

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
BSC32_FLAGS=/nologo /o"$(OUTDIR)/stepconv.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)/stepconv.pdb" /machine:I386 /out:"$(OUTDIR)/stepconv.dll"\
 /implib:"$(OUTDIR)/stepconv.lib" 
LINK32_OBJS= \
	"$(INTDIR)/WVISENSI.OBJ" \
	"$(INTDIR)/WVISCCNV.OBJ" \
	"$(INTDIR)/WSTENSI.OBJ" \
	"$(INTDIR)/WDENSCNV.OBJ" \
	"$(INTDIR)/WDENENSI.OBJ" \
	"$(INTDIR)/VPENSI.OBJ" \
	"$(INTDIR)/VPCONV.OBJ" \
	"$(INTDIR)/TEMPENSI.OBJ" \
	"$(INTDIR)/TEMPCNV.OBJ" \
	"$(INTDIR)/RIENSI.OBJ" \
	"$(INTDIR)/RICONV.OBJ" \
	"$(INTDIR)/PRESSCNV.OBJ" \
	"$(INTDIR)/PRESENSI.OBJ" \
	"$(INTDIR)/NBPENSI.OBJ" \
	"$(INTDIR)/NBPCONV.OBJ" \
	"$(INTDIR)/MWENSI.OBJ" \
	"$(INTDIR)/MWCONV.OBJ" \
	"$(INTDIR)/MVOTENSI.OBJ" \
	"$(INTDIR)/MVOPTCNV.OBJ" \
	"$(INTDIR)/MVNBPCNV.OBJ" \
	"$(INTDIR)/MVBPENSI.OBJ" \
	"$(INTDIR)/LDIFFCNV.OBJ" \
	"$(INTDIR)/LDIFENSI.OBJ" \
	"$(INTDIR)/LDENSCNV.OBJ" \
	"$(INTDIR)/LDENENSI.OBJ" \
	"$(INTDIR)/KOWENSI.OBJ" \
	"$(INTDIR)/KOWCONV.OBJ" \
	"$(INTDIR)/HCENSI.OBJ" \
	"$(INTDIR)/HCCONV.OBJ" \
	"$(INTDIR)/H2OSTCNV.OBJ" \
	"$(INTDIR)/GDIFFCNV.OBJ" \
	"$(INTDIR)/GDIFENSI.OBJ" \
	"$(INTDIR)/AVISENSI.OBJ" \
	"$(INTDIR)/AVISCCNV.OBJ" \
	"$(INTDIR)/AQSENSI.OBJ" \
	"$(INTDIR)/AQSCONV.OBJ" \
	"$(INTDIR)/ADENSCNV.OBJ" \
	"$(INTDIR)/ADENENSI.OBJ" \
	"$(INTDIR)/ACENSI.OBJ" \
	"$(INTDIR)/ACCONV.OBJ"

"$(OUTDIR)\stepconv.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "stepconv - Win32 Debug"

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

ALL : "$(OUTDIR)\stepconv.dll"

CLEAN : 
	-@erase ".\Debug\stepconv.dll"
	-@erase ".\Debug\WVISENSI.OBJ"
	-@erase ".\Debug\WVISCCNV.OBJ"
	-@erase ".\Debug\WSTENSI.OBJ"
	-@erase ".\Debug\WDENSCNV.OBJ"
	-@erase ".\Debug\WDENENSI.OBJ"
	-@erase ".\Debug\VPENSI.OBJ"
	-@erase ".\Debug\VPCONV.OBJ"
	-@erase ".\Debug\TEMPENSI.OBJ"
	-@erase ".\Debug\TEMPCNV.OBJ"
	-@erase ".\Debug\RIENSI.OBJ"
	-@erase ".\Debug\RICONV.OBJ"
	-@erase ".\Debug\PRESSCNV.OBJ"
	-@erase ".\Debug\PRESENSI.OBJ"
	-@erase ".\Debug\NBPENSI.OBJ"
	-@erase ".\Debug\NBPCONV.OBJ"
	-@erase ".\Debug\MWENSI.OBJ"
	-@erase ".\Debug\MWCONV.OBJ"
	-@erase ".\Debug\MVOTENSI.OBJ"
	-@erase ".\Debug\MVOPTCNV.OBJ"
	-@erase ".\Debug\MVNBPCNV.OBJ"
	-@erase ".\Debug\MVBPENSI.OBJ"
	-@erase ".\Debug\LDIFFCNV.OBJ"
	-@erase ".\Debug\LDIFENSI.OBJ"
	-@erase ".\Debug\LDENSCNV.OBJ"
	-@erase ".\Debug\LDENENSI.OBJ"
	-@erase ".\Debug\KOWENSI.OBJ"
	-@erase ".\Debug\KOWCONV.OBJ"
	-@erase ".\Debug\HCENSI.OBJ"
	-@erase ".\Debug\HCCONV.OBJ"
	-@erase ".\Debug\H2OSTCNV.OBJ"
	-@erase ".\Debug\GDIFFCNV.OBJ"
	-@erase ".\Debug\GDIFENSI.OBJ"
	-@erase ".\Debug\AVISENSI.OBJ"
	-@erase ".\Debug\AVISCCNV.OBJ"
	-@erase ".\Debug\AQSENSI.OBJ"
	-@erase ".\Debug\AQSCONV.OBJ"
	-@erase ".\Debug\ADENSCNV.OBJ"
	-@erase ".\Debug\ADENENSI.OBJ"
	-@erase ".\Debug\ACENSI.OBJ"
	-@erase ".\Debug\ACCONV.OBJ"
	-@erase ".\Debug\stepconv.ilk"
	-@erase ".\Debug\stepconv.lib"
	-@erase ".\Debug\stepconv.exp"
	-@erase ".\Debug\stepconv.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "Debug/" /c /nologo /MT
# ADD F90 /Zi /I "Debug/" /c /nologo /MT
F90_PROJ=/Zi /I "Debug/" /c /nologo /MT /Fo"Debug/" /Fd"Debug/stepconv.pdb" 
F90_OBJS=.\Debug/
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/stepconv.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)/stepconv.pdb" /debug /machine:I386\
 /out:"$(OUTDIR)/stepconv.dll" /implib:"$(OUTDIR)/stepconv.lib" 
LINK32_OBJS= \
	"$(INTDIR)/WVISENSI.OBJ" \
	"$(INTDIR)/WVISCCNV.OBJ" \
	"$(INTDIR)/WSTENSI.OBJ" \
	"$(INTDIR)/WDENSCNV.OBJ" \
	"$(INTDIR)/WDENENSI.OBJ" \
	"$(INTDIR)/VPENSI.OBJ" \
	"$(INTDIR)/VPCONV.OBJ" \
	"$(INTDIR)/TEMPENSI.OBJ" \
	"$(INTDIR)/TEMPCNV.OBJ" \
	"$(INTDIR)/RIENSI.OBJ" \
	"$(INTDIR)/RICONV.OBJ" \
	"$(INTDIR)/PRESSCNV.OBJ" \
	"$(INTDIR)/PRESENSI.OBJ" \
	"$(INTDIR)/NBPENSI.OBJ" \
	"$(INTDIR)/NBPCONV.OBJ" \
	"$(INTDIR)/MWENSI.OBJ" \
	"$(INTDIR)/MWCONV.OBJ" \
	"$(INTDIR)/MVOTENSI.OBJ" \
	"$(INTDIR)/MVOPTCNV.OBJ" \
	"$(INTDIR)/MVNBPCNV.OBJ" \
	"$(INTDIR)/MVBPENSI.OBJ" \
	"$(INTDIR)/LDIFFCNV.OBJ" \
	"$(INTDIR)/LDIFENSI.OBJ" \
	"$(INTDIR)/LDENSCNV.OBJ" \
	"$(INTDIR)/LDENENSI.OBJ" \
	"$(INTDIR)/KOWENSI.OBJ" \
	"$(INTDIR)/KOWCONV.OBJ" \
	"$(INTDIR)/HCENSI.OBJ" \
	"$(INTDIR)/HCCONV.OBJ" \
	"$(INTDIR)/H2OSTCNV.OBJ" \
	"$(INTDIR)/GDIFFCNV.OBJ" \
	"$(INTDIR)/GDIFENSI.OBJ" \
	"$(INTDIR)/AVISENSI.OBJ" \
	"$(INTDIR)/AVISCCNV.OBJ" \
	"$(INTDIR)/AQSENSI.OBJ" \
	"$(INTDIR)/AQSCONV.OBJ" \
	"$(INTDIR)/ADENSCNV.OBJ" \
	"$(INTDIR)/ADENENSI.OBJ" \
	"$(INTDIR)/ACENSI.OBJ" \
	"$(INTDIR)/ACCONV.OBJ"

"$(OUTDIR)\stepconv.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
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

# Name "stepconv - Win32 Release"
# Name "stepconv - Win32 Debug"

!IF  "$(CFG)" == "stepconv - Win32 Release"

!ELSEIF  "$(CFG)" == "stepconv - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\WVISENSI.F90

"$(INTDIR)\WVISENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\WVISCCNV.F90

"$(INTDIR)\WVISCCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\WSTENSI.F90

"$(INTDIR)\WSTENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\WDENSCNV.F90

"$(INTDIR)\WDENSCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\WDENENSI.F90

"$(INTDIR)\WDENENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VPENSI.F90

"$(INTDIR)\VPENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VPCONV.F90

"$(INTDIR)\VPCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\TEMPENSI.F90

"$(INTDIR)\TEMPENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\TEMPCNV.F90

"$(INTDIR)\TEMPCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\RIENSI.F90

"$(INTDIR)\RIENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\RICONV.F90

"$(INTDIR)\RICONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PRESSCNV.F90

"$(INTDIR)\PRESSCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PRESENSI.F90

"$(INTDIR)\PRESENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\NBPENSI.F90

"$(INTDIR)\NBPENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\NBPCONV.F90

"$(INTDIR)\NBPCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MWENSI.F90

"$(INTDIR)\MWENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MWCONV.F90

"$(INTDIR)\MWCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MVOTENSI.F90

"$(INTDIR)\MVOTENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MVOPTCNV.F90

"$(INTDIR)\MVOPTCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MVNBPCNV.F90

"$(INTDIR)\MVNBPCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MVBPENSI.F90

"$(INTDIR)\MVBPENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDIFFCNV.F90

"$(INTDIR)\LDIFFCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDIFENSI.F90

"$(INTDIR)\LDIFENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDENSCNV.F90

"$(INTDIR)\LDENSCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDENENSI.F90

"$(INTDIR)\LDENENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\KOWENSI.F90

"$(INTDIR)\KOWENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\KOWCONV.F90

"$(INTDIR)\KOWCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\HCENSI.F90

"$(INTDIR)\HCENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\HCCONV.F90

"$(INTDIR)\HCCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\H2OSTCNV.F90

"$(INTDIR)\H2OSTCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GDIFFCNV.F90

"$(INTDIR)\GDIFFCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\GDIFENSI.F90

"$(INTDIR)\GDIFENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AVISENSI.F90

"$(INTDIR)\AVISENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AVISCCNV.F90

"$(INTDIR)\AVISCCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AQSENSI.F90

"$(INTDIR)\AQSENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AQSCONV.F90

"$(INTDIR)\AQSCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ADENSCNV.F90

"$(INTDIR)\ADENSCNV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ADENENSI.F90

"$(INTDIR)\ADENENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ACENSI.F90

"$(INTDIR)\ACENSI.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ACCONV.F90

"$(INTDIR)\ACCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
# End Target
# End Project
################################################################################
