!CC********************************************************************
!CC
!CC                             MWCONV
!CC          CONVERT MOLECULAR WEIGHT UNITS FROM (-) TO (-).
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for molecular weight.  Right now, the units
!CC               are dimensionless in both cases so there is no
!CC               conversion performed.  However, the routine is
!CC               included in case we are manipulating different
!CC               units in the future.
!CC
!CC Output Variables:
!CC    MWENG =    molecular weight (-)
!CC
!CC Input Variables:
!CC    MWSI =     molecular weight (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE MWCONV(MWENG,MWSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MWCONV
!MS$ ATTRIBUTES ALIAS:'_MWCONV'::MWCONV
!MS$ ATTRIBUTES REFERENCE::MWENG,MWSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION MWENG, MWSI

         MWENG = MWSI * 1.0D0                 

      END

!CC********************************************************************


