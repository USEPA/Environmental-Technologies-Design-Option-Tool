!CC********************************************************************
!CC
!CC                             WDENSCNV
!CC            WATER DENSITY UNITS FROM Kg/m3 TO LBm/Ft3
!CC
!CC Description:  This SUBROUTINE will convert Water Density from
!CC               units of Kg/m3 to units of LBm/Ft3.
!CC
!CC Output Variables:
!CC    WDENG =    Water Density (LBm/Ft3)
!CC
!CC Input Variables:
!CC    WDSI =     Water Density (Kg/m3)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE WDENSCNV(WDENG,WDSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WDENSCNV
!MS$ ATTRIBUTES ALIAS:'_WDENSCNV'::WDENSCNV
!MS$ ATTRIBUTES REFERENCE::WDENG,WDSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WDENG, WDSI

         WDENG = WDSI * 2.20462D0/35.3145D0  

      END

!CC********************************************************************


