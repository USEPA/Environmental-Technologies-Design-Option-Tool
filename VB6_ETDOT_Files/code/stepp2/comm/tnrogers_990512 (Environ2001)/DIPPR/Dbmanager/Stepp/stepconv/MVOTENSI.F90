!CC*******************************************************************
!CC
!CC                                MVOTENSI
!CC               CONVERT MOLAR VOLUME AT OPERATING TEMP FROM LBm/Ft3
!CC               Kg/m3
!CC
!CC Description:  This SUBROUTINE will convert molar volume at operating
!CC               temp from units of LBm/Ft3 to Kg/m3
!CC
!CC Output Variables:
!CC    MVOSI =     Molar Volume at Operating Temp (Kg/m3)
!CC
!CC Input Variables:
!CC    MVOENG =    Molar Volume at Operating Temp (LBm/Ft3)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE MVOTENSI(MVOSI,MVOENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVOTENSI
!MS$ ATTRIBUTES ALIAS:'_MVOTENSI'::MVOTENSI
!MS$ ATTRIBUTES REFERENCE::MVOSI,MVOENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION MVOENG, MVOSI
        MVOSI = MVOENG * 2.20462D0 / 35.3145D0              
      END
 
!CC*******************************************************************


       
