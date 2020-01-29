!CC********************************************************************
!CC
!CC                             AQSCONV
!CC          CONVERT AQUEOUS SOLUBILITY UNITS FROM (-) TO (-).
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for aqueous solubility.  Right now, the units
!CC               are dimensionless in both cases so there is no
!CC               conversion performed.  However, the routine is
!CC               included in case we are manipulating different
!CC               units in the future.
!CC
!CC Output Variables:
!CC    ASENG =    aqueous solubility (-)
!CC
!CC Input Variables:
!CC    ASSI =     aqueous solubility (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE AQSCONV(ASENG,ASSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AQSCONV
!MS$ ATTRIBUTES ALIAS:'_AQSCONV'::AQSCONV
!MS$ ATTRIBUTES REFERENCE::ASENG,ASSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ASENG, ASSI

         ASENG = ASSI * 1.0D0                 

      END

!CC********************************************************************


