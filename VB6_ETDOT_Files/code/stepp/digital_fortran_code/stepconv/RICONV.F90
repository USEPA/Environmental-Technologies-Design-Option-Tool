!CC********************************************************************
!CC
!CC                             RICONV
!CC          CONVERT REFRACTIVE INDEX UNITS FROM (-) TO (-).
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for refractive index.  Right now, the units
!CC               are dimensionless in both cases so there is no
!CC               conversion performed.  However, the routine is
!CC               included in case we are manipulating different
!CC               units in the future.
!CC
!CC Output Variables:
!CC    RIENG =    refractive index (-)
!CC
!CC Input Variables:
!CC    RISI =     refractive index (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE RICONV(RIENG,RISI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::RICONV
!MS$ ATTRIBUTES ALIAS:'_RICONV'::RICONV
!MS$ ATTRIBUTES REFERENCE::RIENG,RISI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION RIENG, RISI

         RIENG = RISI * 1.0D0                 

      END

!CC********************************************************************


