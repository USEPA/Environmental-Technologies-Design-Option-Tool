# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Console Application" 0x0103

!IF "$(CFG)" == ""
CFG=psdm12 - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to psdm12 - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "psdm12 - Win32 Release" && "$(CFG)" != "psdm12 - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "psdm12.mak" CFG="psdm12 - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "psdm12 - Win32 Release" (based on "Win32 (x86) Console Application")
!MESSAGE "psdm12 - Win32 Debug" (based on "Win32 (x86) Console Application")
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
# PROP Target_Last_Scanned "psdm12 - Win32 Debug"
RSC=rc.exe
F90=fl32.exe

!IF  "$(CFG)" == "psdm12 - Win32 Release"

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

ALL : "$(OUTDIR)\psdm12.exe"

CLEAN : 
	-@erase ".\Release\psdm12.exe"
	-@erase ".\Release\FRONT.OBJ"
	-@erase ".\Release\psdm12.obj"
	-@erase ".\Release\DIFFUN.OBJ"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Ox /I "Release/" /c /nologo
# ADD F90 /Ox /I "Release/" /c /nologo
F90_PROJ=/Ox /I "Release/" /c /nologo /Fo"Release/" 
F90_OBJS=.\Release/
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/psdm12.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:console /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:console /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:console /incremental:no\
 /pdb:"$(OUTDIR)/psdm12.pdb" /machine:I386 /out:"$(OUTDIR)/psdm12.exe" 
LINK32_OBJS= \
	"$(INTDIR)/FRONT.OBJ" \
	"$(INTDIR)/psdm12.obj" \
	"$(INTDIR)/DIFFUN.OBJ" \
	".\imsllib.lib"

"$(OUTDIR)\psdm12.exe" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "psdm12 - Win32 Debug"

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

ALL : "$(OUTDIR)\psdm12.exe"

CLEAN : 
	-@erase ".\Debug\psdm12.exe"
	-@erase ".\Debug\psdm12.obj"
	-@erase ".\Debug\FRONT.OBJ"
	-@erase ".\Debug\DIFFUN.OBJ"
	-@erase ".\Debug\psdm12.ilk"
	-@erase ".\Debug\psdm12.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "Debug/" /c /nologo
# ADD F90 /Zi /I "Debug/" /c /nologo
F90_PROJ=/Zi /I "Debug/" /c /nologo /Fo"Debug/" /Fd"Debug/psdm12.pdb" 
F90_OBJS=.\Debug/
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/psdm12.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:console /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:console /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:console /incremental:yes\
 /pdb:"$(OUTDIR)/psdm12.pdb" /debug /machine:I386 /out:"$(OUTDIR)/psdm12.exe" 
LINK32_OBJS= \
	"$(INTDIR)/psdm12.obj" \
	"$(INTDIR)/FRONT.OBJ" \
	"$(INTDIR)/DIFFUN.OBJ" \
	".\imsllib.lib"

"$(OUTDIR)\psdm12.exe" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
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

# Name "psdm12 - Win32 Release"
# Name "psdm12 - Win32 Debug"

!IF  "$(CFG)" == "psdm12 - Win32 Release"

!ELSEIF  "$(CFG)" == "psdm12 - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\psdm12.for
DEP_F90_PSDM1=\
	".\COMMON.FI"\
	

"$(INTDIR)\psdm12.obj" : $(SOURCE) $(DEP_F90_PSDM1) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\FRONT.FOR
DEP_F90_FRONT=\
	".\COMMON.FI"\
	

"$(INTDIR)\FRONT.OBJ" : $(SOURCE) $(DEP_F90_FRONT) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\COMMON.FI

!IF  "$(CFG)" == "psdm12 - Win32 Release"

!ELSEIF  "$(CFG)" == "psdm12 - Win32 Debug"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\imsllib.lib

!IF  "$(CFG)" == "psdm12 - Win32 Release"

!ELSEIF  "$(CFG)" == "psdm12 - Win32 Debug"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\DIFFUN.FOR
DEP_F90_DIFFU=\
	".\COMMON.FI"\
	

"$(INTDIR)\DIFFUN.OBJ" : $(SOURCE) $(DEP_F90_DIFFU) "$(INTDIR)"


# End Source File
# End Target
# End Project
################################################################################
