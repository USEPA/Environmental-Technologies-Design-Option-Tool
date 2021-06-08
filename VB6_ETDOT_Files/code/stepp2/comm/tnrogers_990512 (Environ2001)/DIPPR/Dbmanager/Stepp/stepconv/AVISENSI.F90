!CC*******************************************************************
!CC
!CC                                AVISENSI
!CC          CONVERT AIR VISCOSITY  FROM LBm/Ft-sec TO Kg/m-sec
!CC
!CC
!CC Description:  This SUBROUTINE will convert air viscosity from
!CC               units of LBm/Ft-sec to Kg/m-sec.
!CC
!CC Output Variables:
!CC    AVSI =     Air Viscosity (Kg/m-sec)
!CC
!CC Input Variables:
!CC    AVENG =    Air Viscosity (LBm/Ft-sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE AVISENSI(AVSI,AVENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AVISENSI
!MS$ ATTRIBUTES ALIAS:'_AVISENSI'::AVISENSI
!MS$ ATTRIBUTES REFERENCE::AVSI,AVENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION AVENG, AVSI
        AVSI = AVENG * 3.2808D0 / 2.20462D0                             
      END
 
!CC*******************************************************************


       
