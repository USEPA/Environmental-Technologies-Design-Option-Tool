!CC*******************************************************************
!CC
!CC                                ADENENSI
!CC               CONVERT AIR DENSITY UNITS FROM LBm/Ft3 TO Kg/m3
!CC
!CC Description:  This SUBROUTINE will convert air density from units
!CC               of LBm/Ft3 to units of Kg/m3
!CC
!CC Output Variables:
!CC    ADSI =     Air Density Kg/m3
!CC
!CC Input Variables:
!CC    ADENG =    Air Density LBm/Ft3
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE ADENENSI(ADSI,ADENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ADENENSI
!MS$ ATTRIBUTES ALIAS:'_ADENENSI'::ADENENSI
!MS$ ATTRIBUTES REFERENCE::ADSI,ADENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ADENG, ADSI
        ADSI = ADENG * 35.3145D0 / 2.20462D0
      END
 
!CC*******************************************************************


       
