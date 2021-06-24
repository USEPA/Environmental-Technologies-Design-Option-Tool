!C********************************************************************
!C
!C                             LDENSCNV
!C            LIQUID DENSITY UNITS FROM Kg/m3 TO LBm/Ft3  
!C
!C Description:  This SUBROUTINE will convert Liquid Density from 
!C               units of Kg/m3 to units of LBm/Ft3.
!C
!C Output Variables:
!C    LDENG =    Liquid Density (LBm/Ft3)
!C
!C Input Variables:
!C    LDSI =     Liquid Density (Kg/m3)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE LDENSCNV(LDENG,LDSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDENSCNV
!MS$ ATTRIBUTES ALIAS:'_LDENSCNV':: LDENSCNV
!MS$ ATTRIBUTES REFERENCE:: LDENG,LDSI

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION LDENG, LDSI

LDENG = LDSI * 2.20462D0/35.3145D0  

END SUBROUTINE

!C********************************************************************
