# Microsoft Developer Studio Generated NMAKE File, Format Version 4.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Static Library" 0x0104

!IF "$(CFG)" == ""
CFG=imsllib - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to imsllib - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "imsllib - Win32 Release" && "$(CFG)" !=\
 "imsllib - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "imsllib.mak" CFG="imsllib - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "imsllib - Win32 Release" (based on "Win32 (x86) Static Library")
!MESSAGE "imsllib - Win32 Debug" (based on "Win32 (x86) Static Library")
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

!IF  "$(CFG)" == "imsllib - Win32 Release"

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

ALL : "$(OUTDIR)\imsllib.lib"

CLEAN : 
	-@erase ".\Release\imsllib.lib"
	-@erase ".\Release\umach.obj"
	-@erase ".\Release\s1anum.obj"
	-@erase ".\Release\n1rty.obj"
	-@erase ".\Release\n1rnof.obj"
	-@erase ".\Release\n1rgb.obj"
	-@erase ".\Release\n1rcd.obj"
	-@erase ".\Release\m1vech.obj"
	-@erase ".\Release\m1ve.obj"
	-@erase ".\Release\iwkin.obj"
	-@erase ".\Release\iwkcin.obj"
	-@erase ".\Release\imach.obj"
	-@erase ".\Release\idamax.obj"
	-@erase ".\Release\icase.obj"
	-@erase ".\Release\iachar.obj"
	-@erase ".\Release\i1x.obj"
	-@erase ".\Release\i1kst.obj"
	-@erase ".\Release\i1krl.obj"
	-@erase ".\Release\i1kqu.obj"
	-@erase ".\Release\i1knr.obj"
	-@erase ".\Release\i1kgt.obj"
	-@erase ".\Release\i1kcqu.obj"
	-@erase ".\Release\i1kcgt.obj"
	-@erase ".\Release\i1kc00.obj"
	-@erase ".\Release\i1erif.obj"
	-@erase ".\Release\i1dx.obj"
	-@erase ".\Release\i1cstr.obj"
	-@erase ".\Release\e1usr.obj"
	-@erase ".\Release\e1ucs.obj"
	-@erase ".\Release\e1str.obj"
	-@erase ".\Release\e1stl.obj"
	-@erase ".\Release\e1sti.obj"
	-@erase ".\Release\e1std.obj"
	-@erase ".\Release\e1psh.obj"
	-@erase ".\Release\e1prt.obj"
	-@erase ".\Release\e1pos.obj"
	-@erase ".\Release\e1pop.obj"
	-@erase ".\Release\e1mes.obj"
	-@erase ".\Release\e1inpl.obj"
	-@erase ".\Release\e1init.obj"
	-@erase ".\Release\dswap.obj"
	-@erase ".\Release\dset.obj"
	-@erase ".\Release\dscal.obj"
	-@erase ".\Release\dnrm2.obj"
	-@erase ".\Release\dnr1rr.obj"
	-@erase ".\Release\dneqnf.obj"
	-@erase ".\Release\dn9qnf.obj"
	-@erase ".\Release\dn8qnf.obj"
	-@erase ".\Release\dn7qnf.obj"
	-@erase ".\Release\dn6qnf.obj"
	-@erase ".\Release\dn5qnf.obj"
	-@erase ".\Release\dn4qnf.obj"
	-@erase ".\Release\dn3qnf.obj"
	-@erase ".\Release\dn2qnf.obj"
	-@erase ".\Release\dmurrv.obj"
	-@erase ".\Release\dmrrrr.obj"
	-@erase ".\Release\dmach.obj"
	-@erase ".\Release\dlinrt.obj"
	-@erase ".\Release\dlinrg.obj"
	-@erase ".\Release\dl2trg.obj"
	-@erase ".\Release\dl2nrg.obj"
	-@erase ".\Release\dl2crg.obj"
	-@erase ".\Release\dger.obj"
	-@erase ".\Release\dgemv.obj"
	-@erase ".\Release\dgear.obj"
	-@erase ".\Release\ddot.obj"
	-@erase ".\Release\dcrgrg.obj"
	-@erase ".\Release\dcopy.obj"
	-@erase ".\Release\daxpy.obj"
	-@erase ".\Release\dasum.obj"
	-@erase ".\Release\c1tic.obj"
	-@erase ".\Release\c1tci.obj"
	-@erase ".\Release\amach.obj"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Ox /I "Release/" /c /nologo
# ADD F90 /Ox /I "Release/" /c /nologo
F90_PROJ=/Ox /I "Release/" /c /nologo /Fo"Release/" 
F90_OBJS=.\Release/
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/imsllib.bsc" 
BSC32_SBRS=
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo
LIB32_FLAGS=/nologo /out:"$(OUTDIR)/imsllib.lib" 
LIB32_OBJS= \
	"$(INTDIR)/umach.obj" \
	"$(INTDIR)/s1anum.obj" \
	"$(INTDIR)/n1rty.obj" \
	"$(INTDIR)/n1rnof.obj" \
	"$(INTDIR)/n1rgb.obj" \
	"$(INTDIR)/n1rcd.obj" \
	"$(INTDIR)/m1vech.obj" \
	"$(INTDIR)/m1ve.obj" \
	"$(INTDIR)/iwkin.obj" \
	"$(INTDIR)/iwkcin.obj" \
	"$(INTDIR)/imach.obj" \
	"$(INTDIR)/idamax.obj" \
	"$(INTDIR)/icase.obj" \
	"$(INTDIR)/iachar.obj" \
	"$(INTDIR)/i1x.obj" \
	"$(INTDIR)/i1kst.obj" \
	"$(INTDIR)/i1krl.obj" \
	"$(INTDIR)/i1kqu.obj" \
	"$(INTDIR)/i1knr.obj" \
	"$(INTDIR)/i1kgt.obj" \
	"$(INTDIR)/i1kcqu.obj" \
	"$(INTDIR)/i1kcgt.obj" \
	"$(INTDIR)/i1kc00.obj" \
	"$(INTDIR)/i1erif.obj" \
	"$(INTDIR)/i1dx.obj" \
	"$(INTDIR)/i1cstr.obj" \
	"$(INTDIR)/e1usr.obj" \
	"$(INTDIR)/e1ucs.obj" \
	"$(INTDIR)/e1str.obj" \
	"$(INTDIR)/e1stl.obj" \
	"$(INTDIR)/e1sti.obj" \
	"$(INTDIR)/e1std.obj" \
	"$(INTDIR)/e1psh.obj" \
	"$(INTDIR)/e1prt.obj" \
	"$(INTDIR)/e1pos.obj" \
	"$(INTDIR)/e1pop.obj" \
	"$(INTDIR)/e1mes.obj" \
	"$(INTDIR)/e1inpl.obj" \
	"$(INTDIR)/e1init.obj" \
	"$(INTDIR)/dswap.obj" \
	"$(INTDIR)/dset.obj" \
	"$(INTDIR)/dscal.obj" \
	"$(INTDIR)/dnrm2.obj" \
	"$(INTDIR)/dnr1rr.obj" \
	"$(INTDIR)/dneqnf.obj" \
	"$(INTDIR)/dn9qnf.obj" \
	"$(INTDIR)/dn8qnf.obj" \
	"$(INTDIR)/dn7qnf.obj" \
	"$(INTDIR)/dn6qnf.obj" \
	"$(INTDIR)/dn5qnf.obj" \
	"$(INTDIR)/dn4qnf.obj" \
	"$(INTDIR)/dn3qnf.obj" \
	"$(INTDIR)/dn2qnf.obj" \
	"$(INTDIR)/dmurrv.obj" \
	"$(INTDIR)/dmrrrr.obj" \
	"$(INTDIR)/dmach.obj" \
	"$(INTDIR)/dlinrt.obj" \
	"$(INTDIR)/dlinrg.obj" \
	"$(INTDIR)/dl2trg.obj" \
	"$(INTDIR)/dl2nrg.obj" \
	"$(INTDIR)/dl2crg.obj" \
	"$(INTDIR)/dger.obj" \
	"$(INTDIR)/dgemv.obj" \
	"$(INTDIR)/dgear.obj" \
	"$(INTDIR)/ddot.obj" \
	"$(INTDIR)/dcrgrg.obj" \
	"$(INTDIR)/dcopy.obj" \
	"$(INTDIR)/daxpy.obj" \
	"$(INTDIR)/dasum.obj" \
	"$(INTDIR)/c1tic.obj" \
	"$(INTDIR)/c1tci.obj" \
	"$(INTDIR)/amach.obj"

"$(OUTDIR)\imsllib.lib" : "$(OUTDIR)" $(DEF_FILE) $(LIB32_OBJS)
    $(LIB32) @<<
  $(LIB32_FLAGS) $(DEF_FLAGS) $(LIB32_OBJS)
<<

!ELSEIF  "$(CFG)" == "imsllib - Win32 Debug"

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

ALL : "$(OUTDIR)\imsllib.lib"

CLEAN : 
	-@erase ".\Debug\imsllib.lib"
	-@erase ".\Debug\umach.obj"
	-@erase ".\Debug\s1anum.obj"
	-@erase ".\Debug\n1rty.obj"
	-@erase ".\Debug\n1rnof.obj"
	-@erase ".\Debug\n1rgb.obj"
	-@erase ".\Debug\n1rcd.obj"
	-@erase ".\Debug\m1vech.obj"
	-@erase ".\Debug\m1ve.obj"
	-@erase ".\Debug\iwkin.obj"
	-@erase ".\Debug\iwkcin.obj"
	-@erase ".\Debug\imach.obj"
	-@erase ".\Debug\idamax.obj"
	-@erase ".\Debug\icase.obj"
	-@erase ".\Debug\iachar.obj"
	-@erase ".\Debug\i1x.obj"
	-@erase ".\Debug\i1kst.obj"
	-@erase ".\Debug\i1krl.obj"
	-@erase ".\Debug\i1kqu.obj"
	-@erase ".\Debug\i1knr.obj"
	-@erase ".\Debug\i1kgt.obj"
	-@erase ".\Debug\i1kcqu.obj"
	-@erase ".\Debug\i1kcgt.obj"
	-@erase ".\Debug\i1kc00.obj"
	-@erase ".\Debug\i1erif.obj"
	-@erase ".\Debug\i1dx.obj"
	-@erase ".\Debug\i1cstr.obj"
	-@erase ".\Debug\e1usr.obj"
	-@erase ".\Debug\e1ucs.obj"
	-@erase ".\Debug\e1str.obj"
	-@erase ".\Debug\e1stl.obj"
	-@erase ".\Debug\e1sti.obj"
	-@erase ".\Debug\e1std.obj"
	-@erase ".\Debug\e1psh.obj"
	-@erase ".\Debug\e1prt.obj"
	-@erase ".\Debug\e1pos.obj"
	-@erase ".\Debug\e1pop.obj"
	-@erase ".\Debug\e1mes.obj"
	-@erase ".\Debug\e1inpl.obj"
	-@erase ".\Debug\e1init.obj"
	-@erase ".\Debug\dswap.obj"
	-@erase ".\Debug\dset.obj"
	-@erase ".\Debug\dscal.obj"
	-@erase ".\Debug\dnrm2.obj"
	-@erase ".\Debug\dnr1rr.obj"
	-@erase ".\Debug\dneqnf.obj"
	-@erase ".\Debug\dn9qnf.obj"
	-@erase ".\Debug\dn8qnf.obj"
	-@erase ".\Debug\dn7qnf.obj"
	-@erase ".\Debug\dn6qnf.obj"
	-@erase ".\Debug\dn5qnf.obj"
	-@erase ".\Debug\dn4qnf.obj"
	-@erase ".\Debug\dn3qnf.obj"
	-@erase ".\Debug\dn2qnf.obj"
	-@erase ".\Debug\dmurrv.obj"
	-@erase ".\Debug\dmrrrr.obj"
	-@erase ".\Debug\dmach.obj"
	-@erase ".\Debug\dlinrt.obj"
	-@erase ".\Debug\dlinrg.obj"
	-@erase ".\Debug\dl2trg.obj"
	-@erase ".\Debug\dl2nrg.obj"
	-@erase ".\Debug\dl2crg.obj"
	-@erase ".\Debug\dger.obj"
	-@erase ".\Debug\dgemv.obj"
	-@erase ".\Debug\dgear.obj"
	-@erase ".\Debug\ddot.obj"
	-@erase ".\Debug\dcrgrg.obj"
	-@erase ".\Debug\dcopy.obj"
	-@erase ".\Debug\daxpy.obj"
	-@erase ".\Debug\dasum.obj"
	-@erase ".\Debug\c1tic.obj"
	-@erase ".\Debug\c1tci.obj"
	-@erase ".\Debug\amach.obj"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE F90 /Z7 /I "Debug/" /c /nologo
# ADD F90 /Z7 /I "Debug/" /c /nologo
F90_PROJ=/Z7 /I "Debug/" /c /nologo /Fo"Debug/" 
F90_OBJS=.\Debug/
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/imsllib.bsc" 
BSC32_SBRS=
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo
LIB32_FLAGS=/nologo /out:"$(OUTDIR)/imsllib.lib" 
LIB32_OBJS= \
	"$(INTDIR)/umach.obj" \
	"$(INTDIR)/s1anum.obj" \
	"$(INTDIR)/n1rty.obj" \
	"$(INTDIR)/n1rnof.obj" \
	"$(INTDIR)/n1rgb.obj" \
	"$(INTDIR)/n1rcd.obj" \
	"$(INTDIR)/m1vech.obj" \
	"$(INTDIR)/m1ve.obj" \
	"$(INTDIR)/iwkin.obj" \
	"$(INTDIR)/iwkcin.obj" \
	"$(INTDIR)/imach.obj" \
	"$(INTDIR)/idamax.obj" \
	"$(INTDIR)/icase.obj" \
	"$(INTDIR)/iachar.obj" \
	"$(INTDIR)/i1x.obj" \
	"$(INTDIR)/i1kst.obj" \
	"$(INTDIR)/i1krl.obj" \
	"$(INTDIR)/i1kqu.obj" \
	"$(INTDIR)/i1knr.obj" \
	"$(INTDIR)/i1kgt.obj" \
	"$(INTDIR)/i1kcqu.obj" \
	"$(INTDIR)/i1kcgt.obj" \
	"$(INTDIR)/i1kc00.obj" \
	"$(INTDIR)/i1erif.obj" \
	"$(INTDIR)/i1dx.obj" \
	"$(INTDIR)/i1cstr.obj" \
	"$(INTDIR)/e1usr.obj" \
	"$(INTDIR)/e1ucs.obj" \
	"$(INTDIR)/e1str.obj" \
	"$(INTDIR)/e1stl.obj" \
	"$(INTDIR)/e1sti.obj" \
	"$(INTDIR)/e1std.obj" \
	"$(INTDIR)/e1psh.obj" \
	"$(INTDIR)/e1prt.obj" \
	"$(INTDIR)/e1pos.obj" \
	"$(INTDIR)/e1pop.obj" \
	"$(INTDIR)/e1mes.obj" \
	"$(INTDIR)/e1inpl.obj" \
	"$(INTDIR)/e1init.obj" \
	"$(INTDIR)/dswap.obj" \
	"$(INTDIR)/dset.obj" \
	"$(INTDIR)/dscal.obj" \
	"$(INTDIR)/dnrm2.obj" \
	"$(INTDIR)/dnr1rr.obj" \
	"$(INTDIR)/dneqnf.obj" \
	"$(INTDIR)/dn9qnf.obj" \
	"$(INTDIR)/dn8qnf.obj" \
	"$(INTDIR)/dn7qnf.obj" \
	"$(INTDIR)/dn6qnf.obj" \
	"$(INTDIR)/dn5qnf.obj" \
	"$(INTDIR)/dn4qnf.obj" \
	"$(INTDIR)/dn3qnf.obj" \
	"$(INTDIR)/dn2qnf.obj" \
	"$(INTDIR)/dmurrv.obj" \
	"$(INTDIR)/dmrrrr.obj" \
	"$(INTDIR)/dmach.obj" \
	"$(INTDIR)/dlinrt.obj" \
	"$(INTDIR)/dlinrg.obj" \
	"$(INTDIR)/dl2trg.obj" \
	"$(INTDIR)/dl2nrg.obj" \
	"$(INTDIR)/dl2crg.obj" \
	"$(INTDIR)/dger.obj" \
	"$(INTDIR)/dgemv.obj" \
	"$(INTDIR)/dgear.obj" \
	"$(INTDIR)/ddot.obj" \
	"$(INTDIR)/dcrgrg.obj" \
	"$(INTDIR)/dcopy.obj" \
	"$(INTDIR)/daxpy.obj" \
	"$(INTDIR)/dasum.obj" \
	"$(INTDIR)/c1tic.obj" \
	"$(INTDIR)/c1tci.obj" \
	"$(INTDIR)/amach.obj"

"$(OUTDIR)\imsllib.lib" : "$(OUTDIR)" $(DEF_FILE) $(LIB32_OBJS)
    $(LIB32) @<<
  $(LIB32_FLAGS) $(DEF_FLAGS) $(LIB32_OBJS)
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

# Name "imsllib - Win32 Release"
# Name "imsllib - Win32 Debug"

!IF  "$(CFG)" == "imsllib - Win32 Release"

!ELSEIF  "$(CFG)" == "imsllib - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\umach.f

"$(INTDIR)\umach.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\s1anum.f

"$(INTDIR)\s1anum.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\n1rty.f

"$(INTDIR)\n1rty.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\n1rnof.f

"$(INTDIR)\n1rnof.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\n1rgb.f

"$(INTDIR)\n1rgb.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\n1rcd.f

"$(INTDIR)\n1rcd.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\m1vech.f

"$(INTDIR)\m1vech.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\m1ve.f

"$(INTDIR)\m1ve.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\iwkin.f

"$(INTDIR)\iwkin.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\iwkcin.f

"$(INTDIR)\iwkcin.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\imach.f

"$(INTDIR)\imach.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\idamax.f

"$(INTDIR)\idamax.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\icase.f

"$(INTDIR)\icase.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\iachar.f

"$(INTDIR)\iachar.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1x.f

"$(INTDIR)\i1x.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1kst.f

"$(INTDIR)\i1kst.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1krl.f

"$(INTDIR)\i1krl.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1kqu.f

"$(INTDIR)\i1kqu.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1knr.f

"$(INTDIR)\i1knr.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1kgt.f

"$(INTDIR)\i1kgt.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1kcqu.f

"$(INTDIR)\i1kcqu.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1kcgt.f

"$(INTDIR)\i1kcgt.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1kc00.f

"$(INTDIR)\i1kc00.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1erif.f

"$(INTDIR)\i1erif.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1dx.f

"$(INTDIR)\i1dx.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\i1cstr.f

"$(INTDIR)\i1cstr.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1usr.f

"$(INTDIR)\e1usr.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1ucs.f

"$(INTDIR)\e1ucs.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1str.f

"$(INTDIR)\e1str.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1stl.f

"$(INTDIR)\e1stl.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1sti.f

"$(INTDIR)\e1sti.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1std.f

"$(INTDIR)\e1std.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1psh.f

"$(INTDIR)\e1psh.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1prt.f

"$(INTDIR)\e1prt.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1pos.f

"$(INTDIR)\e1pos.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1pop.f

"$(INTDIR)\e1pop.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1mes.f

"$(INTDIR)\e1mes.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1inpl.f

"$(INTDIR)\e1inpl.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\e1init.f

"$(INTDIR)\e1init.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dswap.f

"$(INTDIR)\dswap.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dset.f

"$(INTDIR)\dset.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dscal.f

"$(INTDIR)\dscal.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dnrm2.f

"$(INTDIR)\dnrm2.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dnr1rr.f

"$(INTDIR)\dnr1rr.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dneqnf.f

"$(INTDIR)\dneqnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn9qnf.f

"$(INTDIR)\dn9qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn8qnf.f

"$(INTDIR)\dn8qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn7qnf.f

"$(INTDIR)\dn7qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn6qnf.f

"$(INTDIR)\dn6qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn5qnf.f

"$(INTDIR)\dn5qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn4qnf.f

"$(INTDIR)\dn4qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn3qnf.f

"$(INTDIR)\dn3qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dn2qnf.f

"$(INTDIR)\dn2qnf.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dmurrv.f

"$(INTDIR)\dmurrv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dmrrrr.f

"$(INTDIR)\dmrrrr.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dmach.f

"$(INTDIR)\dmach.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dlinrt.f

"$(INTDIR)\dlinrt.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dlinrg.f

"$(INTDIR)\dlinrg.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dl2trg.f

"$(INTDIR)\dl2trg.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dl2nrg.f

"$(INTDIR)\dl2nrg.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dl2crg.f

"$(INTDIR)\dl2crg.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dger.f

"$(INTDIR)\dger.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dgemv.f

"$(INTDIR)\dgemv.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dgear.f

"$(INTDIR)\dgear.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\ddot.f

"$(INTDIR)\ddot.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dcrgrg.f

"$(INTDIR)\dcrgrg.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dcopy.f

"$(INTDIR)\dcopy.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\daxpy.f

"$(INTDIR)\daxpy.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\dasum.f

"$(INTDIR)\dasum.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\c1tic.f

"$(INTDIR)\c1tic.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\c1tci.f

"$(INTDIR)\c1tci.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
################################################################################
# Begin Source File

SOURCE=.\amach.f

"$(INTDIR)\amach.obj" : $(SOURCE) "$(INTDIR)"


# End Source File
# End Target
# End Project
################################################################################
