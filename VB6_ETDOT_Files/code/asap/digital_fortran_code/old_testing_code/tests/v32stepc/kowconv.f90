!C********************************************************************
!C
!C                             KOWCONV
!C          CONVERT OCTANOL WATER PARTITION COEF UNITS FROM (-) TO (-).
!C
!C Description:  This SUBROUTINE will handle the conversion of units
!C               for octanol water partition coef.  Right now, the 
!C               units are dimensionless in both cases so there is no
!C               conversion performed.  However, the routine is 
!C               included in case we are manipulating different
!C               units in the future.
!C
!C Output Variables:
!C    OWENG =    octanol water partition coef (-)
!C
!C Input Variables:
!C    OWSI =     octanol water partition coef (-)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE KOWCONV(OWENG,OWSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KOWCONV
!MS$ ATTRIBUTES ALIAS:'_KOWCONV':: KOWCONV
!MS$ ATTRIBUTES REFERENCE::OWENG,OWSI

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION OWENG, OWSI

OWENG = OWSI * 1.0D0                 

END SUBROUTINE

!C********************************************************************
