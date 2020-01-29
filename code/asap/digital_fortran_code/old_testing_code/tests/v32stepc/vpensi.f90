!C*******************************************************************
!C
!C                               VPENSI 
!C             CONVERT VAPOR PRESSURE UNITS FROM psi TO PA
!C
!C Description:  This SUBROUTINE will convert vapor pressure from units      
!C               psi to units of PA.
!C
!C Output Variables:
!C    VPSI =     Vapor Pressure (PA)
!C
!C Input Variables:
!C    VPENG =    Vapor Pressure (psi)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE VPENSI(VPSI,VPENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VPENSI
!MS$ ATTRIBUTES ALIAS:'_VPENSI':: VPENSI
!MS$ ATTRIBUTES REFERENCE::VPSI,VPENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION VPENG, VPSI  

VPSI = VPENG *1.01325D+05 / 14.696D0

END SUBROUTINE
 
!C*******************************************************************


       
