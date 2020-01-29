!C*******************************************************************
!C
!C                               MWENSI 
!C             CONVERT MOLECULAR WEIGHT UNITS FROM (-) TO (-)  
!C
!C Description:  This SUBROUTINE will handle the conversion of units        
!C               for molecular weight.  Right now, the units are
!C               dimensionless in both cases so there is no conversion
!C               performed.  However, the routine is included in case 
!C               we are manipulating different units in the future.                                  
!C
!C Output Variables:
!C    MWSI =     Molecular Weight (-)
!C
!C Input Variables:
!C    MWENG =    Molecular Weight (-)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE MWENSI(MWSI,MWENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MWENSI
!MS$ ATTRIBUTES ALIAS:'_MWENSI':: MWENSI
!MS$ ATTRIBUTES REFERENCE::MWSI,MWENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION MWENG, MWSI  

MWSI = MWENG * 1.0D0                   

END SUBROUTINE

!C*******************************************************************



