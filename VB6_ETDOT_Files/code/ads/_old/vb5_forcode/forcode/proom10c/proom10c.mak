# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Console Application" 0x0103

!IF "$(CFG)" == ""
CFG=proom10c - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to proom10c - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "proom10c - Win32 Release" && "$(CFG)" !=\
 "proom10c - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "proom10c.mak" CFG="proom10c - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "proom10c - Win32 Release" (based on\
 "Win32 (x86) Console Application")
!MESSAGE "proom10c - Win32 Debug" (based on "Win32 (x86) Console Application")
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
# PROP Target_Last_Scanned "proom10c - Win32 Debug"
RSC=rc.exe
F90=fl32.exe

!IF  "$(CFG)" == "proom10c - Win32 Release"

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

ALL : "$(OUTDIR)\proom10c.exe"

CLEAN : 
	-@erase ".\Release\proom10c.exe"
	-@erase ".\Release\dgear.obj"
	-@erase ".\Release\psdm10b.obj"
	-@erase ".\Release\diffun.obj"
	-@erase ".\Release\front.obj"

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
BSC32_FLAGS=/nologo /o"$(OUTDIR)/proom10c.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:console /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:console /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:console /incremental:no\
 /pdb:"$(OUTDIR)/proom10c.pdb" /machine:I386 /out:"$(OUTDIR)/proom10c.exe" 
LINK32_OBJS= \
	"$(INTDIR)/dgear.obj" \
	"$(INTDIR)/psdm10b.obj" \
	"$(INTDIR)/diffun.obj" \
	"$(INTDIR)/front.obj" \
	".\imsllib.lib"

"$(OUTDIR)\proom10c.exe" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "proom10c - Win32 Debug"

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

ALL : "$(OUTDIR)\proom10c.exe"

CLEAN : 
	-@erase ".\Debug\proom10c.exe"
	-@erase ".\Debug\dgear.obj"
	-@erase ".\Debug\diffun.obj"
	-@erase ".\Debug\front.obj"
	-@erase ".\Debug\psdm10b.obj"
	-@erase ".\Debug\proom10c.ilk"
	-@erase ".\Debug\proom10c.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "Debug/" /c /nologo
# ADD F90 /Zi /I "Debug/" /c /nologo
F90_PROJ=/Zi /I "Debug/" /c /nologo /Fo"Debug/" /Fd"Debug/proom10c.pdb" 
F90_OBJS=.\Debug/
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/proom10c.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:console /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:console /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:console /incremental:yes\
 /pdb:"$(OUTDIR)/proom10c.pdb" /debug /machine:I386\
 /out:"$(OUTDIR)/proom10c.exe" 
LINK32_OBJS= \
	"$(INTDIR)/dgear.obj" \
	"$(INTDIR)/diffun.obj" \
	"$(INTDIR)/front.obj" \
	"$(INTDIR)/psdm10b.obj" \
	".\imsllib.lib"

"$(OUTDIR)\proom10c.exe" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
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

# Name "proom10c - Win32 Release"
# Name "proom10c - Win32 Debug"

!IF  "$(CFG)" == "proom10c - Win32 Release"

!ELSEIF  "$(CFG)" == "proom10c - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\psdm10b.for
DEP_F90_PSDM1=\
	".\COMMON.FI"\
	

"$(INTDIR)\psdm10b.obj" : $(SOURCE) $(DEP_F90_PSDM1) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\front.for
DEP_F90_FRONT=\
	".\COMMON.FI"\
	

"$(INTDIR)\front.obj" : $(SOURCE) $(DEP_F90_FRONT) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\diffun.for

"$(INTDIR)\diffun.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dgear.for

"$(INTDIR)\dgear.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\imsllib.lib

!IF  "$(CFG)" == "proom10c - Win32 Release"

!ELSEIF  "$(CFG)" == "proom10c - Win32 Debug"

!ENDIF 

# End Source File
# End Target
# End Project
################################################################################
