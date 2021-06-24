!CC********************************************************************
!CC
!CC                             H2OSTCNV
!CC            WATER SURFACE TENSION UNITS FROM N/m To LBf/Ft
!CC
!CC Description:  This SUBROUTINE will convert water surface tension from
!CC               units of N/m to units of LBf/Ft.
!CC
!CC Output Variables:
!CC    WSTENG =    Water Surface Tension (LBf/Ft)
!CC
!CC Input Variables:
!CC    WSTSI =     Water Surface Tension (N/m)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE H2OSTCNV(WSTENG,WSTSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::H2OSTCNV
!MS$ ATTRIBUTES ALIAS:'_H2OSTCNV'::H2OSTCNV
!MS$ ATTRIBUTES REFERENCE::WSTENG,WSTSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WSTENG, WSTSI

         WSTENG = WSTSI * 0.22481D0/3.2808D0   

      END

!CC********************************************************************


