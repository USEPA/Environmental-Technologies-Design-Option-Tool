!C*************************************************************
!CC
!CC                    DIFLHL
!CC
!CC Description:  This subroutine will calculate liquid diffusivity
!CC               for compounds with molecular weight < 1000 and molar
!CC               volumes between 0.015 and 0.5 m^3/kmol.  It uses
!CC               the Hayduk and Laudie correlation.
!CC
!CC Output Variable:
!CC    DIFL =     Liquid diffusivity (m^2/sec) of compound
!CC
!CC Input Variables:
!CC    VL =       Liquid viscosity (kg/m/sec)
!CC    VB =       Molar volume of compound (m^3/kmol)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C*************************************************************

      SUBROUTINE DIFLHL(DIFL,VL,VB)
!C  ATTRIBUTES DLLEXPORT, STDCALL::DIFLHL
!C  ATTRIBUTES ALIAS:'_DIFLHL@12':: DIFLHL
!C  ATTRIBUTES REFERENCE::DIFL,VL,VB

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFL,VL,VB

         DIFL = (1.326D-4)/(((VL*1000.0D0)**1.14)*((VB*1000.0D0)**0.589))/(100.0D0**2)            

      END

!C*************************************************************

