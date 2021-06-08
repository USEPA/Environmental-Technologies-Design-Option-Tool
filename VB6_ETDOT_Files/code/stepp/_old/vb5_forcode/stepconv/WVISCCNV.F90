!CC********************************************************************
!CC
!CC                             WVISCCNV
!CC            WATER VISCOSITY UNITS FROM Kg/m-sec To LBm/Ft-sec
!CC
!CC Description:  This SUBROUTINE will convert Water Viscosity from
!CC               units of Kg/m-sec to units of LBm/Ft-sec.
!CC
!CC Output Variables:
!CC    WVENG =    Water Viscosity (LBm/Ft-sec)
!CC
!CC Input Variables:
!CC    WVSI =     Water Viscosity (Kg/m-sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE WVISCCNV(WVENG,WVSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WVISCCNV
!MS$ ATTRIBUTES ALIAS:'_WVISCCNV'::WVISCCNV
!MS$ ATTRIBUTES REFERENCE::WVENG,WVSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WVENG, WVSI

         WVENG = WVSI * 2.20462D0/3.2808D0   

      END

!CC********************************************************************


