!CC*******************************************************************
!CC
!CC                                WDENENSI
!CC               CONVERT WATER DENSITY UNITS FROM LBm/Ft3 TO Kg/m3
!CC
!CC Description:  This SUBROUTINE will convert water density from units
!CC               of LBm/Ft3 to units of Kg/m3
!CC
!CC Output Variables:
!CC    WDSI =     Water Density Kg/m3
!CC
!CC Input Variables:
!CC    WDENG =    Water Density LBm/Ft3
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE WDENENSI(WDSI,WDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WDENENSI
!MS$ ATTRIBUTES ALIAS:'_WDENENSI'::WDENENSI
!MS$ ATTRIBUTES REFERENCE::WDSI,WDENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WDENG, WDSI
        WDSI = WDENG * 35.3145D0 / 2.20462D0
      END
 
!CC*******************************************************************


       
