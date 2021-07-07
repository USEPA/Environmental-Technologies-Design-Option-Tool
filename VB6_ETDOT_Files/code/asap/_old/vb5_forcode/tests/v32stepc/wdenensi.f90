!C*******************************************************************
!C
!C                                WDENENSI
!C               CONVERT WATER DENSITY UNITS FROM LBm/Ft3 TO Kg/m3
!C
!C Description:  This SUBROUTINE will convert water density from units 
!C               of LBm/Ft3 to units of Kg/m3
!C
!C Output Variables:
!C    WDSI =     Water Density Kg/m3
!C
!C Input Variables:
!C    WDENG =    Water Density LBm/Ft3
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE WDENENSI(WDSI,WDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WDENENSI
!MS$ ATTRIBUTES ALIAS:'_WDENENSI':: WDENENSI
!MS$ ATTRIBUTES REFERENCE::WDSI,WDENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WDENG, WDSI

WDSI = WDENG * 35.3145D0 / 2.20462D0

END SUBROUTINE
 
!C*******************************************************************


       
