!CC********************************************************************
!CC
!CC                             NBPCONV
!CC          CONVERT NORMAL BOILING POINT UNITS FROM C TO F
!CC
!CC Description:  This SUBROUTINE will convert normal boiling point
!CC               from units of C to units of F.
!CC
!CC Output Variables:
!CC    NBPENG =    Normal Boiling Point (F)
!CC
!CC Input Variables:
!CC    NBPSI =     Normal Boiling Point (C)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE NBPCONV(NBPENG,NBPSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::NBPCONV
!MS$ ATTRIBUTES ALIAS:'_NBPCONV'::NBPCONV
!MS$ ATTRIBUTES REFERENCE::NBPENG,NBPSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION NBPENG, NBPSI

         NBPENG = NBPSI * (9.0D0/5.0D0) + 32.0D0         

      END

!CC********************************************************************


