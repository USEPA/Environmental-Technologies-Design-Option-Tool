!CC********************************************************************
!CC
!CC                             LDENSCNV
!CC            LIQUID DENSITY UNITS FROM Kg/m3 TO LBm/Ft3
!CC
!CC Description:  This SUBROUTINE will convert Liquid Density from
!CC               units of Kg/m3 to units of LBm/Ft3.
!CC
!CC Output Variables:
!CC    LDENG =    Liquid Density (LBm/Ft3)
!CC
!CC Input Variables:
!CC    LDSI =     Liquid Density (Kg/m3)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE LDENSCNV(LDENG,LDSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDENSCNV
!MS$ ATTRIBUTES ALIAS:'_LDENSCNV'::LDENSCNV
!MS$ ATTRIBUTES REFERENCE::LDENG,LDSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION LDENG, LDSI

         LDENG = LDSI * 2.20462D0/35.3145D0  

      END

!CC********************************************************************


