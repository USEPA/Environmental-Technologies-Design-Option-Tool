!C*******************************************************************
!C
!C                                WVISENSI
!C          CONVERT WATER VISCOSITY  FROM LBm/Ft-sec TO Kg/m-sec
!C                    
!C
!C Description:  This SUBROUTINE will convert water viscosity from           
!C               units of LBm/Ft-sec to Kg/m-sec.
!C
!C Output Variables:
!C    WVSI =     Water Viscosity (Kg/m-sec)                   
!C
!C Input Variables:
!C    WVENG =    Water Viscosity (LBm/Ft-sec)                   
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE WVISENSI(WVSI,WVENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WVISENSI
!MS$ ATTRIBUTES ALIAS:'_WVISENSI':: WVISENSI
!MS$ ATTRIBUTES REFERENCE::WVSI,WVENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WVENG, WVSI

WVSI = WVENG * 3.2808D0 / 2.20462D0                             

END	SUBROUTINE

!C*******************************************************************


       
