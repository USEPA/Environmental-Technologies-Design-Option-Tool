!CC********************************************************************
!CC
!CC                             PRESENSI
!CC             CONVERT PRESSURE UNITS FROM psi TO Pa
!CC
!CC Description:  This SUBROUTINE will convert pressure from
!CC               units of psi to units of Pa.
!CC
!CC Output Variables:
!CC    PRESSSI =     Vapor Pressure (Pa)
!CC
!CC Input Variables:
!CC    PRESSENG =    Vapor Pressure (psi)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/21/94)
!CC
!CC********************************************************************

      SUBROUTINE PRESENSI(PRESSSI,PRESSENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PRESENSI
!MS$ ATTRIBUTES ALIAS:'_PRESENSI'::PRESENSI
!MS$ ATTRIBUTES REFERENCE::PRESSSI,PRESSENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION PRESSENG, PRESSSI

         PRESSSI = PRESSENG * 1.01325D+05 / 14.696D0

      END

!CC********************************************************************


