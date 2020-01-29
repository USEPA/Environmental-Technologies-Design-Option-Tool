# Microsoft Developer Studio Project File - Name="imsllib" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Static Library" 0x0104

CFG=imsllib - Win32 Release
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "imsllib.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "imsllib.mak" CFG="imsllib - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "imsllib - Win32 Release" (based on "Win32 (x86) Static Library")
!MESSAGE "imsllib - Win32 Debug" (based on "Win32 (x86) Static Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
F90=df.exe
RSC=rc.exe

!IF  "$(CFG)" == "imsllib - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir ".\Release"
# PROP BASE Intermediate_Dir ".\Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ".\Release"
# PROP Intermediate_Dir ".\Release"
# PROP Target_Dir ""
# ADD BASE F90 /compile_only /include:".\Release/" /nologo /I "Release/"
# ADD F90 /compile_only /include:".\Release/" /include:"Release/" /math_library:fast /nologo
# ADD CPP /FD
# ADD BASE RSC /l 0x409
# ADD RSC /l 0x409
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ELSEIF  "$(CFG)" == "imsllib - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir ".\Debug"
# PROP BASE Intermediate_Dir ".\Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ".\Debug"
# PROP Intermediate_Dir ".\Debug"
# PROP Target_Dir ""
# ADD BASE F90 /compile_only /debug:full /include:".\Debug/" /nologo /I "Debug/"
# ADD F90 /compile_only /debug:full /include:".\Debug/" /include:"Debug/" /nologo /nopdbfile /optimize:0
# ADD CPP /FD
# ADD BASE RSC /l 0x409
# ADD RSC /l 0x409
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ENDIF 

# Begin Target

# Name "imsllib - Win32 Release"
# Name "imsllib - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat;for;f90"
# Begin Source File

SOURCE=.\amach.f
# End Source File
# Begin Source File

SOURCE=.\c1tci.f
# End Source File
# Begin Source File

SOURCE=.\c1tic.f
# End Source File
# Begin Source File

SOURCE=.\dasum.f
# End Source File
# Begin Source File

SOURCE=.\daxpy.f
# End Source File
# Begin Source File

SOURCE=.\dcopy.f
# End Source File
# Begin Source File

SOURCE=.\dcrgrg.f
# End Source File
# Begin Source File

SOURCE=.\ddot.f
# End Source File
# Begin Source File

SOURCE=.\dgear.f
# End Source File
# Begin Source File

SOURCE=.\dgemv.f
# End Source File
# Begin Source File

SOURCE=.\dger.f
# End Source File
# Begin Source File

SOURCE=.\dl2crg.f
# End Source File
# Begin Source File

SOURCE=.\dl2nrg.f
# End Source File
# Begin Source File

SOURCE=.\dl2trg.f
# End Source File
# Begin Source File

SOURCE=.\dlinrg.f
# End Source File
# Begin Source File

SOURCE=.\dlinrt.f
# End Source File
# Begin Source File

SOURCE=.\dmach.f
# End Source File
# Begin Source File

SOURCE=.\dmrrrr.f
# End Source File
# Begin Source File

SOURCE=.\dmurrv.f
# End Source File
# Begin Source File

SOURCE=.\dn2qnf.f
# End Source File
# Begin Source File

SOURCE=.\dn3qnf.f
# End Source File
# Begin Source File

SOURCE=.\dn4qnf.f
# End Source File
# Begin Source File

SOURCE=.\dn5qnf.f
# End Source File
# Begin Source File

SOURCE=.\dn6qnf.f
# End Source File
# Begin Source File

SOURCE=.\dn7qnf.f
# End Source File
# Begin Source File

SOURCE=.\dn8qnf.f
# End Source File
# Begin Source File

SOURCE=.\dn9qnf.f
# End Source File
# Begin Source File

SOURCE=.\dneqnf.f
# End Source File
# Begin Source File

SOURCE=.\dnr1rr.f
# End Source File
# Begin Source File

SOURCE=.\dnrm2.f
# End Source File
# Begin Source File

SOURCE=.\dscal.f
# End Source File
# Begin Source File

SOURCE=.\dset.f
# End Source File
# Begin Source File

SOURCE=.\dswap.f
# End Source File
# Begin Source File

SOURCE=.\e1init.f
# End Source File
# Begin Source File

SOURCE=.\e1inpl.f
# End Source File
# Begin Source File

SOURCE=.\e1mes.f
# End Source File
# Begin Source File

SOURCE=.\e1pop.f
# End Source File
# Begin Source File

SOURCE=.\e1pos.f
# End Source File
# Begin Source File

SOURCE=.\e1prt.f
# End Source File
# Begin Source File

SOURCE=.\e1psh.f
# End Source File
# Begin Source File

SOURCE=.\e1std.f
# End Source File
# Begin Source File

SOURCE=.\e1sti.f
# End Source File
# Begin Source File

SOURCE=.\e1stl.f
# End Source File
# Begin Source File

SOURCE=.\e1str.f
# End Source File
# Begin Source File

SOURCE=.\e1ucs.f
# End Source File
# Begin Source File

SOURCE=.\e1usr.f
# End Source File
# Begin Source File

SOURCE=.\i1cstr.f
# End Source File
# Begin Source File

SOURCE=.\i1dx.f
# End Source File
# Begin Source File

SOURCE=.\i1erif.f
# End Source File
# Begin Source File

SOURCE=.\i1kc00.f
# End Source File
# Begin Source File

SOURCE=.\i1kcgt.f
# End Source File
# Begin Source File

SOURCE=.\i1kcqu.f
# End Source File
# Begin Source File

SOURCE=.\i1kgt.f
# End Source File
# Begin Source File

SOURCE=.\i1knr.f
# End Source File
# Begin Source File

SOURCE=.\i1kqu.f
# End Source File
# Begin Source File

SOURCE=.\i1krl.f
# End Source File
# Begin Source File

SOURCE=.\i1kst.f
# End Source File
# Begin Source File

SOURCE=.\i1x.f
# End Source File
# Begin Source File

SOURCE=.\iachar.f
# End Source File
# Begin Source File

SOURCE=.\icase.f
# End Source File
# Begin Source File

SOURCE=.\idamax.f
# End Source File
# Begin Source File

SOURCE=.\imach.f
# End Source File
# Begin Source File

SOURCE=.\iwkcin.f
# End Source File
# Begin Source File

SOURCE=.\iwkin.f
# End Source File
# Begin Source File

SOURCE=.\m1ve.f
# End Source File
# Begin Source File

SOURCE=.\m1vech.f
# End Source File
# Begin Source File

SOURCE=.\n1rcd.f
# End Source File
# Begin Source File

SOURCE=.\n1rgb.f
# End Source File
# Begin Source File

SOURCE=.\n1rnof.f
# End Source File
# Begin Source File

SOURCE=.\n1rty.f
# End Source File
# Begin Source File

SOURCE=.\s1anum.f
# End Source File
# Begin Source File

SOURCE=.\umach.f
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl;fi;fd"
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;cnt;rtf;gif;jpg;jpeg;jpe"
# End Group
# End Target
# End Project
