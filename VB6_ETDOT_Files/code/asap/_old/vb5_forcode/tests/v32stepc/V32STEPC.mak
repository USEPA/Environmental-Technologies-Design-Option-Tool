# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=V32STEPC - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to V32STEPC - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "V32STEPC - Win32 Release" && "$(CFG)" !=\
 "V32STEPC - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "V32STEPC.mak" CFG="V32STEPC - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "V32STEPC - Win32 Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "V32STEPC - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
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
# PROP Target_Last_Scanned "V32STEPC - Win32 Debug"
RSC=rc.exe
MTL=mktyplib.exe
F90=fl32.exe

!IF  "$(CFG)" == "V32STEPC - Win32 Release"

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

ALL : "$(OUTDIR)\V32STEPC.dll"

CLEAN : 
	-@erase ".\Release\V32STEPC.dll"
	-@erase ".\Release\wvisensi.obj"
	-@erase ".\Release\wvisccnv.obj"
	-@erase ".\Release\wstensi.obj"
	-@erase ".\Release\wdenscnv.obj"
	-@erase ".\Release\wdenensi.obj"
	-@erase ".\Release\vpensi.obj"
	-@erase ".\Release\vpconv.obj"
	-@erase ".\Release\tempensi.obj"
	-@erase ".\Release\tempcnv.obj"
	-@erase ".\Release\riensi.obj"
	-@erase ".\Release\riconv.obj"
	-@erase ".\Release\presscnv.obj"
	-@erase ".\Release\presensi.obj"
	-@erase ".\Release\nbpensi.obj"
	-@erase ".\Release\nbpconv.obj"
	-@erase ".\Release\mwensi.obj"
	-@erase ".\Release\mwconv.obj"
	-@erase ".\Release\mvotensi.obj"
	-@erase ".\Release\mvoptcnv.obj"
	-@erase ".\Release\mvnbpcnv.obj"
	-@erase ".\Release\mvbpensi.obj"
	-@erase ".\Release\ldiffcnv.obj"
	-@erase ".\Release\ldifensi.obj"
	-@erase ".\Release\ldenscnv.obj"
	-@erase ".\Release\ldenensi.obj"
	-@erase ".\Release\kowensi.obj"
	-@erase ".\Release\kowconv.obj"
	-@erase ".\Release\hcensi.obj"
	-@erase ".\Release\hcconv.obj"
	-@erase ".\Release\h2ostcnv.obj"
	-@erase ".\Release\gdiffcnv.obj"
	-@erase ".\Release\gdifensi.obj"
	-@erase ".\Release\avisensi.obj"
	-@erase ".\Release\avisccnv.obj"
	-@erase ".\Release\aqsensi.obj"
	-@erase ".\Release\aqsconv.obj"
	-@erase ".\Release\adenscnv.obj"
	-@erase ".\Release\adenensi.obj"
	-@erase ".\Release\acensi.obj"
	-@erase ".\Release\acconv.obj"
	-@erase ".\Release\V32STEPC.lib"
	-@erase ".\Release\V32STEPC.exp"

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
BSC32_FLAGS=/nologo /o"$(OUTDIR)/V32STEPC.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)/V32STEPC.pdb" /machine:I386 /out:"$(OUTDIR)/V32STEPC.dll"\
 /implib:"$(OUTDIR)/V32STEPC.lib" 
LINK32_OBJS= \
	"$(INTDIR)/wvisensi.obj" \
	"$(INTDIR)/wvisccnv.obj" \
	"$(INTDIR)/wstensi.obj" \
	"$(INTDIR)/wdenscnv.obj" \
	"$(INTDIR)/wdenensi.obj" \
	"$(INTDIR)/vpensi.obj" \
	"$(INTDIR)/vpconv.obj" \
	"$(INTDIR)/tempensi.obj" \
	"$(INTDIR)/tempcnv.obj" \
	"$(INTDIR)/riensi.obj" \
	"$(INTDIR)/riconv.obj" \
	"$(INTDIR)/presscnv.obj" \
	"$(INTDIR)/presensi.obj" \
	"$(INTDIR)/nbpensi.obj" \
	"$(INTDIR)/nbpconv.obj" \
	"$(INTDIR)/mwensi.obj" \
	"$(INTDIR)/mwconv.obj" \
	"$(INTDIR)/mvotensi.obj" \
	"$(INTDIR)/mvoptcnv.obj" \
	"$(INTDIR)/mvnbpcnv.obj" \
	"$(INTDIR)/mvbpensi.obj" \
	"$(INTDIR)/ldiffcnv.obj" \
	"$(INTDIR)/ldifensi.obj" \
	"$(INTDIR)/ldenscnv.obj" \
	"$(INTDIR)/ldenensi.obj" \
	"$(INTDIR)/kowensi.obj" \
	"$(INTDIR)/kowconv.obj" \
	"$(INTDIR)/hcensi.obj" \
	"$(INTDIR)/hcconv.obj" \
	"$(INTDIR)/h2ostcnv.obj" \
	"$(INTDIR)/gdiffcnv.obj" \
	"$(INTDIR)/gdifensi.obj" \
	"$(INTDIR)/avisensi.obj" \
	"$(INTDIR)/avisccnv.obj" \
	"$(INTDIR)/aqsensi.obj" \
	"$(INTDIR)/aqsconv.obj" \
	"$(INTDIR)/adenscnv.obj" \
	"$(INTDIR)/adenensi.obj" \
	"$(INTDIR)/acensi.obj" \
	"$(INTDIR)/acconv.obj"

"$(OUTDIR)\V32STEPC.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "V32STEPC - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "V32STEPC"
# PROP BASE Intermediate_Dir "V32STEPC"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "V32STEPC"
# PROP Intermediate_Dir "V32STEPC"
# PROP Target_Dir ""
OUTDIR=.\V32STEPC
INTDIR=.\V32STEPC

ALL : "$(OUTDIR)\V32STEPC.dll"

CLEAN : 
	-@erase ".\V32STEPC\V32STEPC.dll"
	-@erase ".\V32STEPC\wvisensi.obj"
	-@erase ".\V32STEPC\wvisccnv.obj"
	-@erase ".\V32STEPC\wstensi.obj"
	-@erase ".\V32STEPC\wdenscnv.obj"
	-@erase ".\V32STEPC\wdenensi.obj"
	-@erase ".\V32STEPC\vpensi.obj"
	-@erase ".\V32STEPC\vpconv.obj"
	-@erase ".\V32STEPC\tempensi.obj"
	-@erase ".\V32STEPC\tempcnv.obj"
	-@erase ".\V32STEPC\riensi.obj"
	-@erase ".\V32STEPC\riconv.obj"
	-@erase ".\V32STEPC\presscnv.obj"
	-@erase ".\V32STEPC\presensi.obj"
	-@erase ".\V32STEPC\nbpensi.obj"
	-@erase ".\V32STEPC\nbpconv.obj"
	-@erase ".\V32STEPC\mwensi.obj"
	-@erase ".\V32STEPC\mwconv.obj"
	-@erase ".\V32STEPC\mvotensi.obj"
	-@erase ".\V32STEPC\mvoptcnv.obj"
	-@erase ".\V32STEPC\mvnbpcnv.obj"
	-@erase ".\V32STEPC\mvbpensi.obj"
	-@erase ".\V32STEPC\ldiffcnv.obj"
	-@erase ".\V32STEPC\ldifensi.obj"
	-@erase ".\V32STEPC\ldenscnv.obj"
	-@erase ".\V32STEPC\ldenensi.obj"
	-@erase ".\V32STEPC\kowensi.obj"
	-@erase ".\V32STEPC\kowconv.obj"
	-@erase ".\V32STEPC\hcensi.obj"
	-@erase ".\V32STEPC\hcconv.obj"
	-@erase ".\V32STEPC\h2ostcnv.obj"
	-@erase ".\V32STEPC\gdiffcnv.obj"
	-@erase ".\V32STEPC\gdifensi.obj"
	-@erase ".\V32STEPC\avisensi.obj"
	-@erase ".\V32STEPC\avisccnv.obj"
	-@erase ".\V32STEPC\aqsensi.obj"
	-@erase ".\V32STEPC\aqsconv.obj"
	-@erase ".\V32STEPC\adenscnv.obj"
	-@erase ".\V32STEPC\adenensi.obj"
	-@erase ".\V32STEPC\acensi.obj"
	-@erase ".\V32STEPC\acconv.obj"
	-@erase ".\V32STEPC\V32STEPC.ilk"
	-@erase ".\V32STEPC\V32STEPC.lib"
	-@erase ".\V32STEPC\V32STEPC.exp"
	-@erase ".\V32STEPC\V32STEPC.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "V32STEPC/" /c /nologo /MT
# ADD F90 /Zi /I "V32STEPC/" /c /nologo /MT
F90_PROJ=/Zi /I "V32STEPC/" /c /nologo /MT /Fo"V32STEPC/"\
 /Fd"V32STEPC/V32STEPC.pdb" 
F90_OBJS=.\V32STEPC/
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/V32STEPC.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)/V32STEPC.pdb" /debug /machine:I386\
 /out:"$(OUTDIR)/V32STEPC.dll" /implib:"$(OUTDIR)/V32STEPC.lib" 
LINK32_OBJS= \
	"$(INTDIR)/wvisensi.obj" \
	"$(INTDIR)/wvisccnv.obj" \
	"$(INTDIR)/wstensi.obj" \
	"$(INTDIR)/wdenscnv.obj" \
	"$(INTDIR)/wdenensi.obj" \
	"$(INTDIR)/vpensi.obj" \
	"$(INTDIR)/vpconv.obj" \
	"$(INTDIR)/tempensi.obj" \
	"$(INTDIR)/tempcnv.obj" \
	"$(INTDIR)/riensi.obj" \
	"$(INTDIR)/riconv.obj" \
	"$(INTDIR)/presscnv.obj" \
	"$(INTDIR)/presensi.obj" \
	"$(INTDIR)/nbpensi.obj" \
	"$(INTDIR)/nbpconv.obj" \
	"$(INTDIR)/mwensi.obj" \
	"$(INTDIR)/mwconv.obj" \
	"$(INTDIR)/mvotensi.obj" \
	"$(INTDIR)/mvoptcnv.obj" \
	"$(INTDIR)/mvnbpcnv.obj" \
	"$(INTDIR)/mvbpensi.obj" \
	"$(INTDIR)/ldiffcnv.obj" \
	"$(INTDIR)/ldifensi.obj" \
	"$(INTDIR)/ldenscnv.obj" \
	"$(INTDIR)/ldenensi.obj" \
	"$(INTDIR)/kowensi.obj" \
	"$(INTDIR)/kowconv.obj" \
	"$(INTDIR)/hcensi.obj" \
	"$(INTDIR)/hcconv.obj" \
	"$(INTDIR)/h2ostcnv.obj" \
	"$(INTDIR)/gdiffcnv.obj" \
	"$(INTDIR)/gdifensi.obj" \
	"$(INTDIR)/avisensi.obj" \
	"$(INTDIR)/avisccnv.obj" \
	"$(INTDIR)/aqsensi.obj" \
	"$(INTDIR)/aqsconv.obj" \
	"$(INTDIR)/adenscnv.obj" \
	"$(INTDIR)/adenensi.obj" \
	"$(INTDIR)/acensi.obj" \
	"$(INTDIR)/acconv.obj"

"$(OUTDIR)\V32STEPC.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
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

# Name "V32STEPC - Win32 Release"
# Name "V32STEPC - Win32 Debug"

!IF  "$(CFG)" == "V32STEPC - Win32 Release"

!ELSEIF  "$(CFG)" == "V32STEPC - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\wvisensi.f90

"$(INTDIR)\wvisensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\wvisccnv.f90

"$(INTDIR)\wvisccnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\wstensi.f90

"$(INTDIR)\wstensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\wdenscnv.f90

"$(INTDIR)\wdenscnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\wdenensi.f90

"$(INTDIR)\wdenensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\vpensi.f90

"$(INTDIR)\vpensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\vpconv.f90

"$(INTDIR)\vpconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\tempensi.f90

"$(INTDIR)\tempensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\tempcnv.f90

"$(INTDIR)\tempcnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\riensi.f90

"$(INTDIR)\riensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\riconv.f90

"$(INTDIR)\riconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\presscnv.f90

"$(INTDIR)\presscnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\presensi.f90

"$(INTDIR)\presensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\nbpensi.f90

"$(INTDIR)\nbpensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\nbpconv.f90

"$(INTDIR)\nbpconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\mwensi.f90

"$(INTDIR)\mwensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\mwconv.f90

"$(INTDIR)\mwconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\mvotensi.f90

"$(INTDIR)\mvotensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\mvoptcnv.f90

"$(INTDIR)\mvoptcnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\mvnbpcnv.f90

"$(INTDIR)\mvnbpcnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\mvbpensi.f90

"$(INTDIR)\mvbpensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ldiffcnv.f90

"$(INTDIR)\ldiffcnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ldifensi.f90

"$(INTDIR)\ldifensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ldenscnv.f90

"$(INTDIR)\ldenscnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ldenensi.f90

"$(INTDIR)\ldenensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\kowensi.f90

"$(INTDIR)\kowensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\kowconv.f90

"$(INTDIR)\kowconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\hcensi.f90

"$(INTDIR)\hcensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\hcconv.f90

"$(INTDIR)\hcconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\h2ostcnv.f90

"$(INTDIR)\h2ostcnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\gdiffcnv.f90

"$(INTDIR)\gdiffcnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\gdifensi.f90

"$(INTDIR)\gdifensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\avisensi.f90

"$(INTDIR)\avisensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\avisccnv.f90

"$(INTDIR)\avisccnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\aqsensi.f90

"$(INTDIR)\aqsensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\aqsconv.f90

"$(INTDIR)\aqsconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\adenscnv.f90

"$(INTDIR)\adenscnv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\adenensi.f90

"$(INTDIR)\adenensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\acensi.f90

"$(INTDIR)\acensi.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\acconv.f90

"$(INTDIR)\acconv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
# End Target
# End Project
################################################################################
