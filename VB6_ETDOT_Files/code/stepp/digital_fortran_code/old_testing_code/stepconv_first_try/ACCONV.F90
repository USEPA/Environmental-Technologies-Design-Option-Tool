!CC********************************************************************
!CC
!CC                             ACCONV
!CC          CONVERT ACTIVITY COEFFICIENT UNITS FROM (-) TO (-).
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for activity coefficient.  Right now, the units
!CC               are dimensionless in both cases so there is no
!CC               conversion performed.  However, the routine is
!CC               included in case we are manipulating different
!CC               units in the future.
!CC
!CC Output Variables:
!CC    ACENG =    Activity Coefficient (-)
!CC
!CC Input Variables:
!CC    ACSI =     Activity Coefficient (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE ACCONV(ACENG,ACSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ACCONV
!MS$ ATTRIBUTES ALIAS:'_ACCONV'::ACCONV
!MS$ ATTRIBUTES REFERENCE::ACENG,ACSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ACENG, ACSI

         ACENG = ACSI * 1.0D0                 

      END

!CC********************************************************************


