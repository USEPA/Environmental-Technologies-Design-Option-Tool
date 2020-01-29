!CC*******************************************************************
!CC
!CC                               VPENSI
!CC             CONVERT VAPOR PRESSURE UNITS FROM psi TO PA
!CC
!CC Description:  This SUBROUTINE will convert vapor pressure from units
!CC               psi to units of PA.
!CC
!CC Output Variables:
!CC    VPSI =     Vapor Pressure (PA)
!CC
!CC Input Variables:
!CC    VPENG =    Vapor Pressure (psi)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE VPENSI(VPSI,VPENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VPENSI
!MS$ ATTRIBUTES ALIAS:'_VPENSI'::VPENSI
!MS$ ATTRIBUTES REFERENCE::VPSI,VPENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION VPENG, VPSI  
        VPSI = VPENG *1.01325D+05 / 14.696D0
      END
 
!CC*******************************************************************


       
