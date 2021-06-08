This fortran code is modified from StEPP version 1.0 
for StEPP version 2.

stepp.dll is compiled in Compaq Visual Fortran Version 6.1.

Dimensioning on Binary Interaction Parameter Database is
changed from original values:

  MA=53
  NA=96

to new values:

  MA=58
  NA=116

in all PARAMETER statements where above parameters are dimensioned.

The routines in stepp.dll that are CHANGED by this are:

   10 Routines MODIFIED

ACCALL.F90
Aqscall.f90
Binpar.f90
Fgrp.f90
Hc1call.f90
Kowcall.f90
Molwt.f90
Mwtcall.f90
Vbbpcall.f90
Vbmsch.f90

The routines in stepp.dll that are NOT MODIFIED by this change are:

31 Routines NOT Modified

Airdens.f90
Airvisc.f90
Aqsfit.f90
Aqsol.f90
Dbdens.f90
Difgwl.f90
Diflhl.f90
Diflpol.f90
Diflwc.f90
Error.f90
Fgrpcall.f90
Getgam.f90
H2odens.f90
H2ost.f90
H2ovisc.f90
Hc2call.f90
Hcdbconv.f90
Henfit.f90
Henry.f90
Initvs.f90
Lddbcall.f90
Ldgccall.f90
Newton.f90
Orgdens.f90
Parms.f90
Partc.f90
Regress.f90
Unimod.f90
Vaporp.f90
Vbmatt.f90
Vprcall.f90

D. Hokanson
7/9/01