!C*******************************************************************
!C
!C                               HCENSI 
!C             CONVERT HENRY'S CONSTANT UNITS FROM (-) TO (-)  
!C
!C Description:  This SUBROUTINE will handle the conversion of units        
!C               for Henry's Constant.  Right now, the units are
!C               dimensionless in both cases so there is no conversion
!C               performed.  However, the routine is included in case 
!C               we are manipulating different units in the future.                                  
!C
!C Output Variables:
!C    HCSI =     Henry's Constant (-)
!C
!C Input Variables:
!C    HCENG =    Henry's Constant (-)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE HCENSI(HCSI,HCENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HCENSI
!MS$ ATTRIBUTES ALIAS:'_HCENSI':: HCENSI
!MS$ ATTRIBUTES REFERENCE::HCSI,HCENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION HCENG, HCSI  

HCSI = HCENG * 1.0D0                   

END	 SUBROUTINE

!C*******************************************************************


       
