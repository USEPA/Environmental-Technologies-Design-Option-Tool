!CC*******************************************************************
!CC
!CC                                WVISENSI
!CC          CONVERT WATER VISCOSITY  FROM LBm/Ft-sec TO Kg/m-sec
!CC
!CC
!CC Description:  This SUBROUTINE will convert water viscosity from
!CC               units of LBm/Ft-sec to Kg/m-sec.
!CC
!CC Output Variables:
!CC    WVSI =     Water Viscosity (Kg/m-sec)
!CC
!CC Input Variables:
!CC    WVENG =    Water Viscosity (LBm/Ft-sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE WVISENSI(WVSI,WVENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WVISENSI
!MS$ ATTRIBUTES ALIAS:'_WVISENSI'::WVISENSI
!MS$ ATTRIBUTES REFERENCE::WVSI,WVENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WVENG, WVSI
        WVSI = WVENG * 3.2808D0 / 2.20462D0                             
      END
 
!CC*******************************************************************


       
