!CC********************************************************************
!CC
!CC                             ADENSCNV
!CC              AIR DENSITY UNITS FROM Kg/m3 TO LBm/Ft3
!CC
!CC Description:  This SUBROUTINE will convert Air Density from
!CC               units of Kg/m3 to units of LBm/Ft3.
!CC
!CC Output Variables:
!CC    ADENG =    Air Density (LBm/Ft3)
!CC
!CC Input Variables:
!CC    ADSI =     Air Density (Kg/m3)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE ADENSCNV(ADENG,ADSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ADENSCNV
!MS$ ATTRIBUTES ALIAS:'_ADENSCNV'::ADENSCNV
!MS$ ATTRIBUTES REFERENCE::ADENG,ADSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ADENG, ADSI

         ADENG = ADSI * 2.20462D0/35.3145D0  

      END

!CC********************************************************************


