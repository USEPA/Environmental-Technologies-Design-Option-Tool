!C*******************************************************************
!C
!C                                AVISENSI
!C          CONVERT AIR VISCOSITY  FROM LBm/Ft-sec TO Kg/m-sec
!C                    
!C
!C Description:  This SUBROUTINE will convert air viscosity from           
!C               units of LBm/Ft-sec to Kg/m-sec.
!C
!C Output Variables:
!C    AVSI =     Air Viscosity (Kg/m-sec)                   
!C
!C Input Variables:
!C    AVENG =    Air Viscosity (LBm/Ft-sec)                   
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE AVISENSI(AVSI,AVENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AVISENSI
!MS$ ATTRIBUTES ALIAS:'_AVISENSI':: AVISENSI
!MS$ ATTRIBUTES REFERENCE::AVSI,AVENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION AVENG, AVSI
        AVSI = AVENG * 3.2808D0 / 2.20462D0                             

END SUBROUTINE

!C*******************************************************************


       
