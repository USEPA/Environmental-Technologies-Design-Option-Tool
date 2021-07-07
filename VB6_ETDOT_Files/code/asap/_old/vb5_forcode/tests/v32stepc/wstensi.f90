!C*******************************************************************
!C
!C                                WSTENSI
!C            CONVERT WATER SURFACE TENSION FROM LBf/Ft to N/m         
!C                    
!C
!C Description:  This SUBROUTINE will convert water surface tension                
!C               from units of LBf/Ft to N/m.
!C
!C Output Variables:
!C    WSTSI =     Water Surface Tension (N/m)                   
!C
!C Input Variables:
!C    WSTENG =    Water Surface Tension (LBf/Ft)                   
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE WSTENSI(WSTSI,WSTENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WSTENSI
!MS$ ATTRIBUTES ALIAS:'_WSTENSI':: WSTENSI
!MS$ ATTRIBUTES REFERENCE::WSTSI,WSTENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WSTENG, WSTSI

WSTSI = WSTENG * 3.2808D0 / 0.22481D0                             

END SUBROUTINE

!C*******************************************************************


       
