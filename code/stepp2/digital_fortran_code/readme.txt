  *** Changes July 12, 2001

Binpar.F90 was changed as described in the email from Chaitanya
to Hokanson below:

From: "Chaitanya" <cbelwal@mtu.edu>
To: "Dave Hokanson" <drhokans@mtu.edu>
Cc: <tnrogers@mtu.edu>
Subject: Request for some changes
Date: Thu, 12 Jul 2001 09:55:13 -0400
X-MSMail-Priority: Normal
X-Mailer: Microsoft Outlook Express 4.72.3110.5
X-MimeOLE: Produced By Microsoft MimeOLE V4.72.3110.3

Hello Dave,

There are a few more changes to be done in the DLL to load another binary
paramater file ( Global Data ). I installed fortran powerstaion but it seems
I have the old version ( version 4 1994-95) , I cant read ur project files
(dsp / dsw ), hence I think it will be best if you can re compile it with
your compiler.

The changes required are in the binpar file. I have edited the orig binpar
file ( from the updated DLL ) and am sending it across with this as file
name binpar_modified.f90. The lines which I have modified / added,  have my
comments as "modified by CB". All changes were done by Dr. Rogers and you
can review it if all is ok.

Thank you for your help.

Regards,
Chaitanya

Attachment Converted: "H:\pc_mail\hokanson\attach\BINPAR_modified.F90"


  *** Changes July 9, 2001

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