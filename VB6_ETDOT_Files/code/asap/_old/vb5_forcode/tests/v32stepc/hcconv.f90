!C********************************************************************
!C
!C                             HCCONV
!C          CONVERT HENRY'S CONSTANT UNITS FROM (-) TO (-).
!C
!C Description:  This SUBROUTINE will handle the conversion of units
!C               for Henry's constant.  Right now, the units
!C               are dimensionless in both cases so there is no
!C               conversion performed.  However, the routine is 
!C               included in case we are manipulating different
!C               units in the future.
!C
!C Output Variables:
!C    HCENG =    Henry's Constant (-)
!C
!C Input Variables:
!C    HCSI =     Henry's Constant (-)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE HCCONV(HCENG,HCSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HCCONV
!MS$ ATTRIBUTES ALIAS:'_HCCONV':: HCCONV
!MS$ ATTRIBUTES REFERENCE::HCENG,HCSI

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION HCENG, HCSI

HCENG = HCSI * 1.0D0                 

END	SUBROUTINE

!C********************************************************************
