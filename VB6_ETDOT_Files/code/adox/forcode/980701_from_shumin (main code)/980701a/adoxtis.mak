# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Console Application" 0x0103

!IF "$(CFG)" == ""
CFG=pwtis - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to pwtis - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "pwtis - Win32 Release" && "$(CFG)" != "pwtis - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "pwtis.mak" CFG="pwtis - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "pwtis - Win32 Release" (based on "Win32 (x86) Console Application")
!MESSAGE "pwtis - Win32 Debug" (based on "Win32 (x86) Console Application")
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
F90=fl32.exe
RSC=rc.exe

!IF  "$(CFG)" == "pwtis - Win32 Release"

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

ALL : "$(OUTDIR)\pwtis.exe"

CLEAN : 
	-@erase ".\Release\pwtis.exe"
	-@erase ".\Release\photorate.obj"
	-@erase ".\Release\odequatn.obj"
	-@erase ".\Release\fcn.obj"
	-@erase ".\Release\diffun.obj"
	-@erase ".\Release\dgear.obj"
	-@erase ".\Release\adoxtis.obj"

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
BSC32_FLAGS=/nologo /o"$(OUTDIR)/pwtis.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:console /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:console /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:console /incremental:no\
 /pdb:"$(OUTDIR)/pwtis.pdb" /machine:I386 /out:"$(OUTDIR)/pwtis.exe" 
LINK32_OBJS= \
	".\Release\photorate.obj" \
	".\Release\odequatn.obj" \
	".\Release\fcn.obj" \
	".\Release\diffun.obj" \
	".\Release\dgear.obj" \
	".\Release\adoxtis.obj" \
	"C:\MSDEV\LIB\MATHS.LIB" \
	"C:\MSDEV\LIB\MATHD.LIB"

"$(OUTDIR)\pwtis.exe" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "pwtis - Win32 Debug"

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

ALL : "$(OUTDIR)\pwtis.exe"

CLEAN : 
	-@erase ".\Debug\pwtis.exe"
	-@erase ".\Debug\photorate.obj"
	-@erase ".\Debug\odequatn.obj"
	-@erase ".\Debug\fcn.obj"
	-@erase ".\Debug\diffun.obj"
	-@erase ".\Debug\dgear.obj"
	-@erase ".\Debug\adoxtis.obj"
	-@erase ".\Debug\pwtis.ilk"
	-@erase ".\Debug\pwtis.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "Debug/" /c /nologo
# ADD F90 /Zi /I "Debug/" /c /nologo
F90_PROJ=/Zi /I "Debug/" /c /nologo /Fo"Debug/" /Fd"Debug/pwtis.pdb" 
F90_OBJS=.\Debug/
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/pwtis.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:console /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:console /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:console /incremental:yes\
 /pdb:"$(OUTDIR)/pwtis.pdb" /debug /machine:I386 /out:"$(OUTDIR)/pwtis.exe" 
LINK32_OBJS= \
	".\Debug\photorate.obj" \
	".\Debug\odequatn.obj" \
	".\Debug\fcn.obj" \
	".\Debug\diffun.obj" \
	".\Debug\dgear.obj" \
	".\Debug\adoxtis.obj" \
	"C:\MSDEV\LIB\MATHS.LIB" \
	"C:\MSDEV\LIB\MATHD.LIB"

"$(OUTDIR)\pwtis.exe" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
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

# Name "pwtis - Win32 Release"
# Name "pwtis - Win32 Debug"

!IF  "$(CFG)" == "pwtis - Win32 Release"

!ELSEIF  "$(CFG)" == "pwtis - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\photorate.f

"$(INTDIR)\photorate.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\odequatn.f

"$(INTDIR)\odequatn.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\model.out

!IF  "$(CFG)" == "pwtis - Win32 Release"

!ELSEIF  "$(CFG)" == "pwtis - Win32 Debug"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\model.dat

!IF  "$(CFG)" == "pwtis - Win32 Release"

!ELSEIF  "$(CFG)" == "pwtis - Win32 Debug"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\fcn.f

"$(INTDIR)\fcn.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\diffun.f

"$(INTDIR)\diffun.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dgear.f

"$(INTDIR)\dgear.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\adoxtis.f

"$(INTDIR)\adoxtis.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\adoxpath.txt

!IF  "$(CFG)" == "pwtis - Win32 Release"

!ELSEIF  "$(CFG)" == "pwtis - Win32 Debug"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=C:\MSDEV\LIB\MATHS.LIB

!IF  "$(CFG)" == "pwtis - Win32 Release"

!ELSEIF  "$(CFG)" == "pwtis - Win32 Debug"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=C:\MSDEV\LIB\MATHD.LIB

!IF  "$(CFG)" == "pwtis - Win32 Release"

!ELSEIF  "$(CFG)" == "pwtis - Win32 Debug"

!ENDIF 

# End Source File
# End Target
# End Project
################################################################################
