!C********************************************************************
!C
!C                             H2OSTCNV
!C            WATER SURFACE TENSION UNITS FROM N/m To LBf/Ft
!C
!C Description:  This SUBROUTINE will convert water surface tension from 
!C               units of N/m to units of LBf/Ft.
!C
!C Output Variables:
!C    WSTENG =    Water Surface Tension (LBf/Ft)
!C
!C Input Variables:
!C    WSTSI =     Water Surface Tension (N/m)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE H2OSTCNV(WSTENG,WSTSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::H2OSTCNV
!MS$ ATTRIBUTES ALIAS:'_H2OSTCNV':: H2OSTCNV
!MS$ ATTRIBUTES REFERENCE::WSTENG,WSTSI

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION WSTENG, WSTSI

WSTENG = WSTSI * 0.22481D0/3.2808D0   

END SUBROUTINE

!C********************************************************************
