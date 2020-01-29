# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=stepp - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to stepp - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "stepp - Win32 Release" && "$(CFG)" != "stepp - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "stepp.mak" CFG="stepp - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "stepp - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "stepp - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
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

!IF  "$(CFG)" == "stepp - Win32 Release"

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

ALL : "$(OUTDIR)\stepp.dll"

CLEAN : 
	-@erase ".\Release\stepp.dll"
	-@erase ".\Release\VPRCALL.OBJ"
	-@erase ".\Release\VBMSCH.OBJ"
	-@erase ".\Release\VBMATT.OBJ"
	-@erase ".\Release\VBBPCALL.OBJ"
	-@erase ".\Release\VAPORP.OBJ"
	-@erase ".\Release\UNIMOD.OBJ"
	-@erase ".\Release\REGRESS.OBJ"
	-@erase ".\Release\PARTC.OBJ"
	-@erase ".\Release\PARMS.OBJ"
	-@erase ".\Release\ORGDENS.OBJ"
	-@erase ".\Release\NEWTON.OBJ"
	-@erase ".\Release\MWTCALL.OBJ"
	-@erase ".\Release\MOLWT.OBJ"
	-@erase ".\Release\LDGCCALL.OBJ"
	-@erase ".\Release\LDDBCALL.OBJ"
	-@erase ".\Release\KOWCALL.OBJ"
	-@erase ".\Release\INITVS.OBJ"
	-@erase ".\Release\HENRY.OBJ"
	-@erase ".\Release\HENFIT.OBJ"
	-@erase ".\Release\HCDBCONV.OBJ"
	-@erase ".\Release\HC2CALL.OBJ"
	-@erase ".\Release\HC1CALL.OBJ"
	-@erase ".\Release\H2OVISC.OBJ"
	-@erase ".\Release\H2OST.OBJ"
	-@erase ".\Release\H2ODENS.OBJ"
	-@erase ".\Release\GETGAM.OBJ"
	-@erase ".\Release\FGRPCALL.OBJ"
	-@erase ".\Release\FGRP.OBJ"
	-@erase ".\Release\ERROR.OBJ"
	-@erase ".\Release\DIFLWC.OBJ"
	-@erase ".\Release\DIFLPOL.OBJ"
	-@erase ".\Release\DIFLHL.OBJ"
	-@erase ".\Release\DIFGWL.OBJ"
	-@erase ".\Release\DBDENS.OBJ"
	-@erase ".\Release\BINPAR.OBJ"
	-@erase ".\Release\AQSOL.OBJ"
	-@erase ".\Release\AQSFIT.OBJ"
	-@erase ".\Release\AQSCALL.OBJ"
	-@erase ".\Release\AIRVISC.OBJ"
	-@erase ".\Release\AIRDENS.OBJ"
	-@erase ".\Release\ACCALL.OBJ"
	-@erase ".\Release\stepp.lib"
	-@erase ".\Release\stepp.exp"

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
BSC32_FLAGS=/nologo /o"$(OUTDIR)/stepp.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)/stepp.pdb" /machine:I386 /out:"$(OUTDIR)/stepp.dll"\
 /implib:"$(OUTDIR)/stepp.lib" 
LINK32_OBJS= \
	"$(INTDIR)/VPRCALL.OBJ" \
	"$(INTDIR)/VBMSCH.OBJ" \
	"$(INTDIR)/VBMATT.OBJ" \
	"$(INTDIR)/VBBPCALL.OBJ" \
	"$(INTDIR)/VAPORP.OBJ" \
	"$(INTDIR)/UNIMOD.OBJ" \
	"$(INTDIR)/REGRESS.OBJ" \
	"$(INTDIR)/PARTC.OBJ" \
	"$(INTDIR)/PARMS.OBJ" \
	"$(INTDIR)/ORGDENS.OBJ" \
	"$(INTDIR)/NEWTON.OBJ" \
	"$(INTDIR)/MWTCALL.OBJ" \
	"$(INTDIR)/MOLWT.OBJ" \
	"$(INTDIR)/LDGCCALL.OBJ" \
	"$(INTDIR)/LDDBCALL.OBJ" \
	"$(INTDIR)/KOWCALL.OBJ" \
	"$(INTDIR)/INITVS.OBJ" \
	"$(INTDIR)/HENRY.OBJ" \
	"$(INTDIR)/HENFIT.OBJ" \
	"$(INTDIR)/HCDBCONV.OBJ" \
	"$(INTDIR)/HC2CALL.OBJ" \
	"$(INTDIR)/HC1CALL.OBJ" \
	"$(INTDIR)/H2OVISC.OBJ" \
	"$(INTDIR)/H2OST.OBJ" \
	"$(INTDIR)/H2ODENS.OBJ" \
	"$(INTDIR)/GETGAM.OBJ" \
	"$(INTDIR)/FGRPCALL.OBJ" \
	"$(INTDIR)/FGRP.OBJ" \
	"$(INTDIR)/ERROR.OBJ" \
	"$(INTDIR)/DIFLWC.OBJ" \
	"$(INTDIR)/DIFLPOL.OBJ" \
	"$(INTDIR)/DIFLHL.OBJ" \
	"$(INTDIR)/DIFGWL.OBJ" \
	"$(INTDIR)/DBDENS.OBJ" \
	"$(INTDIR)/BINPAR.OBJ" \
	"$(INTDIR)/AQSOL.OBJ" \
	"$(INTDIR)/AQSFIT.OBJ" \
	"$(INTDIR)/AQSCALL.OBJ" \
	"$(INTDIR)/AIRVISC.OBJ" \
	"$(INTDIR)/AIRDENS.OBJ" \
	"$(INTDIR)/ACCALL.OBJ"

"$(OUTDIR)\stepp.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

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

ALL : "$(OUTDIR)\stepp.dll"

CLEAN : 
	-@erase ".\Debug\stepp.dll"
	-@erase ".\Debug\VPRCALL.OBJ"
	-@erase ".\Debug\VBMSCH.OBJ"
	-@erase ".\Debug\VBMATT.OBJ"
	-@erase ".\Debug\VBBPCALL.OBJ"
	-@erase ".\Debug\VAPORP.OBJ"
	-@erase ".\Debug\UNIMOD.OBJ"
	-@erase ".\Debug\REGRESS.OBJ"
	-@erase ".\Debug\PARTC.OBJ"
	-@erase ".\Debug\PARMS.OBJ"
	-@erase ".\Debug\ORGDENS.OBJ"
	-@erase ".\Debug\NEWTON.OBJ"
	-@erase ".\Debug\MWTCALL.OBJ"
	-@erase ".\Debug\MOLWT.OBJ"
	-@erase ".\Debug\LDGCCALL.OBJ"
	-@erase ".\Debug\LDDBCALL.OBJ"
	-@erase ".\Debug\KOWCALL.OBJ"
	-@erase ".\Debug\INITVS.OBJ"
	-@erase ".\Debug\HENRY.OBJ"
	-@erase ".\Debug\HENFIT.OBJ"
	-@erase ".\Debug\HCDBCONV.OBJ"
	-@erase ".\Debug\HC2CALL.OBJ"
	-@erase ".\Debug\HC1CALL.OBJ"
	-@erase ".\Debug\H2OVISC.OBJ"
	-@erase ".\Debug\H2OST.OBJ"
	-@erase ".\Debug\H2ODENS.OBJ"
	-@erase ".\Debug\GETGAM.OBJ"
	-@erase ".\Debug\FGRPCALL.OBJ"
	-@erase ".\Debug\FGRP.OBJ"
	-@erase ".\Debug\ERROR.OBJ"
	-@erase ".\Debug\DIFLWC.OBJ"
	-@erase ".\Debug\DIFLPOL.OBJ"
	-@erase ".\Debug\DIFLHL.OBJ"
	-@erase ".\Debug\DIFGWL.OBJ"
	-@erase ".\Debug\DBDENS.OBJ"
	-@erase ".\Debug\BINPAR.OBJ"
	-@erase ".\Debug\AQSOL.OBJ"
	-@erase ".\Debug\AQSFIT.OBJ"
	-@erase ".\Debug\AQSCALL.OBJ"
	-@erase ".\Debug\AIRVISC.OBJ"
	-@erase ".\Debug\AIRDENS.OBJ"
	-@erase ".\Debug\ACCALL.OBJ"
	-@erase ".\Debug\stepp.ilk"
	-@erase ".\Debug\stepp.lib"
	-@erase ".\Debug\stepp.exp"
	-@erase ".\Debug\stepp.pdb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Zi /I "Debug/" /c /nologo /MT
# ADD F90 /Zi /I "Debug/" /c /nologo /MT
F90_PROJ=/Zi /I "Debug/" /c /nologo /MT /Fo"Debug/" /Fd"Debug/stepp.pdb" 
F90_OBJS=.\Debug/
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/stepp.bsc" 
BSC32_SBRS=
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib /nologo /subsystem:windows /dll /debug /machine:I386
LINK32_FLAGS=kernel32.lib /nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)/stepp.pdb" /debug /machine:I386 /out:"$(OUTDIR)/stepp.dll"\
 /implib:"$(OUTDIR)/stepp.lib" 
LINK32_OBJS= \
	"$(INTDIR)/VPRCALL.OBJ" \
	"$(INTDIR)/VBMSCH.OBJ" \
	"$(INTDIR)/VBMATT.OBJ" \
	"$(INTDIR)/VBBPCALL.OBJ" \
	"$(INTDIR)/VAPORP.OBJ" \
	"$(INTDIR)/UNIMOD.OBJ" \
	"$(INTDIR)/REGRESS.OBJ" \
	"$(INTDIR)/PARTC.OBJ" \
	"$(INTDIR)/PARMS.OBJ" \
	"$(INTDIR)/ORGDENS.OBJ" \
	"$(INTDIR)/NEWTON.OBJ" \
	"$(INTDIR)/MWTCALL.OBJ" \
	"$(INTDIR)/MOLWT.OBJ" \
	"$(INTDIR)/LDGCCALL.OBJ" \
	"$(INTDIR)/LDDBCALL.OBJ" \
	"$(INTDIR)/KOWCALL.OBJ" \
	"$(INTDIR)/INITVS.OBJ" \
	"$(INTDIR)/HENRY.OBJ" \
	"$(INTDIR)/HENFIT.OBJ" \
	"$(INTDIR)/HCDBCONV.OBJ" \
	"$(INTDIR)/HC2CALL.OBJ" \
	"$(INTDIR)/HC1CALL.OBJ" \
	"$(INTDIR)/H2OVISC.OBJ" \
	"$(INTDIR)/H2OST.OBJ" \
	"$(INTDIR)/H2ODENS.OBJ" \
	"$(INTDIR)/GETGAM.OBJ" \
	"$(INTDIR)/FGRPCALL.OBJ" \
	"$(INTDIR)/FGRP.OBJ" \
	"$(INTDIR)/ERROR.OBJ" \
	"$(INTDIR)/DIFLWC.OBJ" \
	"$(INTDIR)/DIFLPOL.OBJ" \
	"$(INTDIR)/DIFLHL.OBJ" \
	"$(INTDIR)/DIFGWL.OBJ" \
	"$(INTDIR)/DBDENS.OBJ" \
	"$(INTDIR)/BINPAR.OBJ" \
	"$(INTDIR)/AQSOL.OBJ" \
	"$(INTDIR)/AQSFIT.OBJ" \
	"$(INTDIR)/AQSCALL.OBJ" \
	"$(INTDIR)/AIRVISC.OBJ" \
	"$(INTDIR)/AIRDENS.OBJ" \
	"$(INTDIR)/ACCALL.OBJ"

"$(OUTDIR)\stepp.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
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

# Name "stepp - Win32 Release"
# Name "stepp - Win32 Debug"

!IF  "$(CFG)" == "stepp - Win32 Release"

!ELSEIF  "$(CFG)" == "stepp - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\VPRCALL.F90

"$(INTDIR)\VPRCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VBMSCH.F90

"$(INTDIR)\VBMSCH.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VBMATT.F90

"$(INTDIR)\VBMATT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VBBPCALL.F90

"$(INTDIR)\VBBPCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\VAPORP.F90

"$(INTDIR)\VAPORP.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\UNIMOD.F90

"$(INTDIR)\UNIMOD.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\REGRESS.F90

"$(INTDIR)\REGRESS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PARTC.F90

"$(INTDIR)\PARTC.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\PARMS.F90

"$(INTDIR)\PARMS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ORGDENS.F90

"$(INTDIR)\ORGDENS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\NEWTON.F90

"$(INTDIR)\NEWTON.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MWTCALL.F90

"$(INTDIR)\MWTCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\MOLWT.F90

"$(INTDIR)\MOLWT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDGCCALL.F90

"$(INTDIR)\LDGCCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\LDDBCALL.F90

"$(INTDIR)\LDDBCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\KOWCALL.F90

"$(INTDIR)\KOWCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\INITVS.F90

"$(INTDIR)\INITVS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\HENRY.F90

"$(INTDIR)\HENRY.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\HENFIT.F90

"$(INTDIR)\HENFIT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\HCDBCONV.F90

"$(INTDIR)\HCDBCONV.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\HC2CALL.F90

"$(INTDIR)\HC2CALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\HC1CALL.F90

"$(INTDIR)\HC1CALL.OBJ" : $(SOURCE) "$(INTDIR)"


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

SOURCE=.\GETGAM.F90

"$(INTDIR)\GETGAM.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\FGRPCALL.F90

"$(INTDIR)\FGRPCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\FGRP.F90

"$(INTDIR)\FGRP.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ERROR.F90

"$(INTDIR)\ERROR.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\DIFLWC.F90

"$(INTDIR)\DIFLWC.OBJ" : $(SOURCE) "$(INTDIR)"


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

SOURCE=.\DBDENS.F90

"$(INTDIR)\DBDENS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\BINPAR.F90

"$(INTDIR)\BINPAR.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AQSOL.F90

"$(INTDIR)\AQSOL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AQSFIT.F90

"$(INTDIR)\AQSFIT.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AQSCALL.F90

"$(INTDIR)\AQSCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AIRVISC.F90

"$(INTDIR)\AIRVISC.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\AIRDENS.F90

"$(INTDIR)\AIRDENS.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ACCALL.F90

"$(INTDIR)\ACCALL.OBJ" : $(SOURCE) "$(INTDIR)"


# End Source File
# End Target
# End Project
################################################################################
