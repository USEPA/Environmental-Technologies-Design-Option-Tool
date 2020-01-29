!C********************************************************************
!C
!C                             PRESENSI
!C             CONVERT PRESSURE UNITS FROM psi TO Pa
!C
!C Description:  This SUBROUTINE will convert pressure from 
!C               units of psi to units of Pa.
!C
!C Output Variables:
!C    PRESSSI =     Vapor Pressure (Pa)    
!C
!C Input Variables:
!C    PRESSENG =    Vapor Pressure (psi)
!C
!C History:
!C    Function written by D. Hokanson (6/21/94)
!C
!C********************************************************************

SUBROUTINE PRESENSI(PRESSSI,PRESSENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PRESENSI
!MS$ ATTRIBUTES ALIAS:'_PRESENSI':: PRESENSI
!MS$ ATTRIBUTES REFERENCE::PRESSSI,PRESSENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION PRESSENG, PRESSSI

PRESSSI = PRESSENG * 1.01325D+05 / 14.696D0

END SUBROUTINE

!C********************************************************************
