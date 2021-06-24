!CC*******************************************************************
!CC
!CC                                WSTENSI
!CC            CONVERT WATER SURFACE TENSION FROM LBf/Ft to N/m
!CC
!CC
!CC Description:  This SUBROUTINE will convert water surface tension
!CC               from units of LBf/Ft to N/m.
!CC
!CC Output Variables:
!CC    WSTSI =     Water Surface Tension (N/m)
!CC
!CC Input Variables:
!CC    WSTENG =    Water Surface Tension (LBf/Ft)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE WSTENSI(WSTSI,WSTENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WSTENSI
!MS$ ATTRIBUTES ALIAS:'_WSTENSI'::WSTENSI
!MS$ ATTRIBUTES REFERENCE::WSTSI,WSTENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WSTENG, WSTSI
        WSTSI = WSTENG * 3.2808D0 / 0.22481D0                             
      END
 
!CC*******************************************************************


       
