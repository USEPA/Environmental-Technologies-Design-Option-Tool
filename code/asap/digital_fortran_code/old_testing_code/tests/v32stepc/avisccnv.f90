!C********************************************************************
!C
!C                             AVISCCNV
!C            AIR VISCOSITY UNITS FROM Kg/m-sec To LBm/Ft-sec
!C
!C Description:  This SUBROUTINE will convert Air Viscosity from 
!C               units of Kg/m-sec to units of LBm/Ft-sec.
!C
!C Output Variables:
!C    AVENG =    Air Viscosity (LBm/Ft-sec)
!C
!C Input Variables:
!C    AVSI =     Air Viscosity (Kg/m-sec)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE AVISCCNV(AVENG,AVSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AVISCCNV
!MS$ ATTRIBUTES ALIAS:'_AVISCCNV':: AVISCCNV
!MS$ ATTRIBUTES REFERENCE::AVENG,AVSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION AVENG, AVSI

         AVENG = AVSI * 2.20462D0/3.2808D0   

END SUBROUTINE

!C********************************************************************
