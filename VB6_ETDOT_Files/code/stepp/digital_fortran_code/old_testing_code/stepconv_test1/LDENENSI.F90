!CC*******************************************************************
!CC
!CC                                LDENENSI
!CC               CONVERT LIQUID DENSITY UNITS FROM LBm/Ft3 TO Kg/m3
!CC
!CC Description:  This SUBROUTINE will convert liquid density from units
!CC               of LBm/Ft3 to units of Kg/m3
!CC
!CC Output Variables:
!CC    LDSI =     Liquid Density Kg/m3
!CC
!CC Input Variables:
!CC    LDENG =    Liquid Density LBm/Ft3
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE LDENENSI(LDSI,LDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDENENSI
!MS$ ATTRIBUTES ALIAS:'_LDENENSI'::LDENENSI
!MS$ ATTRIBUTES REFERENCE::LDSI,LDENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION LDENG, LDSI
        LDSI = LDENG * 35.3145D0 / 2.20462D0
      END
 
!CC*******************************************************************


       
