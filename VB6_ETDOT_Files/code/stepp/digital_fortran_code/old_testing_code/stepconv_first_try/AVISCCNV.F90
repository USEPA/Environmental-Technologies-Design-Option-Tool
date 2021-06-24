!CC********************************************************************
!CC
!CC                             AVISCCNV
!CC            AIR VISCOSITY UNITS FROM Kg/m-sec To LBm/Ft-sec
!CC
!CC Description:  This SUBROUTINE will convert Air Viscosity from
!CC               units of Kg/m-sec to units of LBm/Ft-sec.
!CC
!CC Output Variables:
!CC    AVENG =    Air Viscosity (LBm/Ft-sec)
!CC
!CC Input Variables:
!CC    AVSI =     Air Viscosity (Kg/m-sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE AVISCCNV(AVENG,AVSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AVISCCNV
!MS$ ATTRIBUTES ALIAS:'_AVISCCNV'::AVISCCNV
!MS$ ATTRIBUTES REFERENCE::AVENG,AVSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION AVENG, AVSI

         AVENG = AVSI * 2.20462D0/3.2808D0   

      END

!CC********************************************************************


