!C*******************************************************************
!C
!C                                LDENENSI
!C               CONVERT LIQUID DENSITY UNITS FROM LBm/Ft3 TO Kg/m3
!C
!C Description:  This SUBROUTINE will convert liquid density from units 
!C               of LBm/Ft3 to units of Kg/m3
!C
!C Output Variables:
!C    LDSI =     Liquid Density Kg/m3
!C
!C Input Variables:
!C    LDENG =    Liquid Density LBm/Ft3
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE LDENENSI(LDSI,LDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDENENSI
!MS$ ATTRIBUTES ALIAS:'_LDENENSI':: LDENENSI
!MS$ ATTRIBUTES REFERENCE::LDSI,LDENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION LDENG, LDSI

LDSI = LDENG * 35.3145D0 / 2.20462D0

END SUBROUTINE
 
!C*******************************************************************


       
