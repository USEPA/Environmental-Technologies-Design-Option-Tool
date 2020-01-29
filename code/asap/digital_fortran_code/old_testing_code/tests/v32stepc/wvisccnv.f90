!C********************************************************************
!C
!C                             WVISCCNV
!C            WATER VISCOSITY UNITS FROM Kg/m-sec To LBm/Ft-sec
!C
!C Description:  This SUBROUTINE will convert Water Viscosity from 
!C               units of Kg/m-sec to units of LBm/Ft-sec.
!C
!C Output Variables:
!C    WVENG =    Water Viscosity (LBm/Ft-sec)
!C
!C Input Variables:
!C    WVSI =     Water Viscosity (Kg/m-sec)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE WVISCCNV(WVENG,WVSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WVISCCNV
!MS$ ATTRIBUTES ALIAS:'_WVISCCNV':: WVISCCNV
!MS$ ATTRIBUTES REFERENCE::WVENG,WVSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WVENG, WVSI

WVENG = WVSI * 2.20462D0/3.2808D0   

END SUBROUTINE

!C********************************************************************
