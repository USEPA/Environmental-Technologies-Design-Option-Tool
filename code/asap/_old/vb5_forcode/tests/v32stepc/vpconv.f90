!C********************************************************************
!C
!C                             VPCONV
!C          CONVERT VAPOR PRESSURE UNITS FROM Pa TO psi
!C
!C Description:  This SUBROUTINE will convert vapor pressure from 
!C               units of Pa to units of psi.
!C
!C Output Variables:
!C    VPENG =    Vapor Pressure (psi)
!C
!C Input Variables:
!C    VPSI =     Vapor Pressure (Pa)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE VPCONV(VPENG,VPSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VPCONV
!MS$ ATTRIBUTES ALIAS:'_VPCONV':: VPCONV
!MS$ ATTRIBUTES REFERENCE::VPENG,VPSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION VPENG, VPSI

VPENG = VPSI * 14.696D0 / 1.01325D+05

END SUBROUTINE

!C********************************************************************
