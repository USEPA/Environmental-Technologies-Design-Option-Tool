!CC********************************************************************
!CC
!CC                             HCCONV
!CC          CONVERT HENRY'S CONSTANT UNITS FROM (-) TO (-).
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for Henry's constant.  Right now, the units
!CC               are dimensionless in both cases so there is no
!CC               conversion performed.  However, the routine is
!CC               included in case we are manipulating different
!CC               units in the future.
!CC
!CC Output Variables:
!CC    HCENG =    Henry's Constant (-)
!CC
!CC Input Variables:
!CC    HCSI =     Henry's Constant (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE HCCONV(HCENG,HCSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HCCONV
!MS$ ATTRIBUTES ALIAS:'_HCCONV'::HCCONV
!MS$ ATTRIBUTES REFERENCE::HCENG,HCSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION HCENG, HCSI

         HCENG = HCSI * 1.0D0                 

      END

!CC********************************************************************


