!C********************************************************************
!C
!C                             MWCONV
!C          CONVERT MOLECULAR WEIGHT UNITS FROM (-) TO (-).
!C
!C Description:  This SUBROUTINE will handle the conversion of units
!C               for molecular weight.  Right now, the units
!C               are dimensionless in both cases so there is no
!C               conversion performed.  However, the routine is 
!C               included in case we are manipulating different
!C               units in the future.
!C
!C Output Variables:
!C    MWENG =    molecular weight (-)
!C
!C Input Variables:
!C    MWSI =     molecular weight (-)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE MWCONV(MWENG,MWSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MWCONV
!MS$ ATTRIBUTES ALIAS:'_MWCONV':: MWCONV
!MS$ ATTRIBUTES REFERENCE::MWENG,MWSI

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION MWENG, MWSI

MWENG = MWSI * 1.0D0                 

END SUBROUTINE

!C********************************************************************
