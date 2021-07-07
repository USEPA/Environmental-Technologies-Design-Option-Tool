!C********************************************************************
!C
!C                             ACCONV
!C          CONVERT ACTIVITY COEFFICIENT UNITS FROM (-) TO (-).
!C
!C Description:  This SUBROUTINE will handle the conversion of units
!C               for activity coefficient.  Right now, the units
!C               are dimensionless in both cases so there is no
!C               conversion performed.  However, the routine is 
!C               included in case we are manipulating different
!C               units in the future.
!C
!C Output Variables:
!C    ACENG =    Activity Coefficient (-)
!C
!C Input Variables:
!C    ACSI =     Activity Coefficient (-)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE ACCONV(ACENG,ACSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ACCONV
!MS$ ATTRIBUTES ALIAS:'_ACCONV':: ACCONV
!MS$ ATTRIBUTES REFERENCE::ACENG,ACSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ACENG, ACSI

         ACENG = ACSI * 1.0D0                 

END SUBROUTINE

!C********************************************************************
