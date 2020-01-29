!CC********************************************************************
!CC
!CC                             KOWCONV
!CC          CONVERT OCTANOL WATER PARTITION COEF UNITS FROM (-) TO (-).
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for octanol water partition coef.  Right now, the
!CC               units are dimensionless in both cases so there is no
!CC               conversion performed.  However, the routine is
!CC               included in case we are manipulating different
!CC               units in the future.
!CC
!CC Output Variables:
!CC    OWENG =    octanol water partition coef (-)
!CC
!CC Input Variables:
!CC    OWSI =     octanol water partition coef (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE KOWCONV(OWENG,OWSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KOWCONV
!MS$ ATTRIBUTES ALIAS:'_KOWCONV'::KOWCONV
!MS$ ATTRIBUTES REFERENCE::OWENG,OWSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION OWENG, OWSI

         OWENG = OWSI * 1.0D0                 

      END

!CC********************************************************************


