!CC*******************************************************************
!CC
!CC                                MVBPENSI
!CC          CONVERT MOLAR VOLUME AT NORMAL BOILING POINT FROM LBm/Ft3
!CC               Kg/m3
!CC
!CC Description:  This SUBROUTINE will convert molar volume at normal
!CC               boiling point from units of LBm/Ft3 to Kg/m3.
!CC
!CC Output Variables:
!CC    MVBSI =     Molar Volume at Normal Boiling Point (Kg/m3)
!CC
!CC Input Variables:
!CC    MVBENG =    Molar Volume at Normal Boiling Point (LBm/Ft3)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE MVBPENSI(MVBSI,MVBENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVBPENSI
!MS$ ATTRIBUTES ALIAS:'_MVBPENSI'::MVBPENSI
!MS$ ATTRIBUTES REFERENCE::MVBSI,MVBENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION MVBENG, MVBSI
        MVBSI = MVBENG * 2.20462D0 / 35.3145D0              
      END
 
!CC*******************************************************************


       
