!CC********************************************************************
!CC
!CC                             VPCONV
!CC          CONVERT VAPOR PRESSURE UNITS FROM Pa TO psi
!CC
!CC Description:  This SUBROUTINE will convert vapor pressure from
!CC               units of Pa to units of psi.
!CC
!CC Output Variables:
!CC    VPENG =    Vapor Pressure (psi)
!CC
!CC Input Variables:
!CC    VPSI =     Vapor Pressure (Pa)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE VPCONV(VPENG,VPSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VPCONV
!MS$ ATTRIBUTES ALIAS:'_VPCONV'::VPCONV
!MS$ ATTRIBUTES REFERENCE::VPENG,VPSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION VPENG, VPSI

         VPENG = VPSI * 14.696D0 / 1.01325D+05

      END

!CC********************************************************************


