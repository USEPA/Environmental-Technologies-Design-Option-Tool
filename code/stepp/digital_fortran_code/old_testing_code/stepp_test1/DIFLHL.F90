!C*************************************************************
!CC
!CC                    DIFLHL
!CC
!CC Description:  This subroutine will calculate liquid diffusivity
!CC               for compounds.  It is generally valid for
!CC               molecular weight < 1000 and molar
!CC               volumes between 0.015 and 0.5 m^3/kmol.  It uses
!CC               the Hayduk and Laudie correlation.
!CC
!CC Output Variables:
!CC    DIFL =     Liquid diffusivity (m^2/sec) of compound
!CC    ERRORF =   Error flag
!CC    SRCSHT =   Source of this value (Short version)
!CC    SRCLNG =   Source of this value (Long version)
!CC    DIFLT =    Temperature of this calculation (C)
!CC
!CC Input Variables:
!CC    VB =       Molar volume at normal boiling point (m^3/kmol)
!CC    TEMPOP =   Temperature of the calculation (C)
!CC    MW =       Molecular weight (kg/kmol)
!CC
!CC Variables Internal to Subroutine DIFLHL:
!CC    VL =       Liquid viscosity (kg/m/sec)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C*************************************************************


      SUBROUTINE DIFLHL(DIFL,VB,TEMPOP,MW,ERRORF,SRCSHT,SRCLNG,DIFLT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DIFLHL
!MS$ ATTRIBUTES ALIAS:'_DIFLHL':: DIFLHL
!MS$ ATTRIBUTES REFERENCE::DIFL,VB,TEMPOP,MW,ERRORF,SRCSHT,SRCLNG,DIFLT

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFL,VL,VB,TEMPOP,MW
      INTEGER ERRORF,SRCSHT,SRCLNG

         ERRORF = 0
         DIFLT = TEMPOP
         SRCSHT = 10

         CALL H2OVISC(VL,TEMPOP,MERR,IDUMSHT,IDUMLNG,DUMTMP)
         DIFL = (1.326D-4)/(((VL*1000.0D0)**1.14)*((VB*1000.0D0)**0.589))/(100.0D0**2)
         IF ((MW.GT.1000).AND.(VB.LT.(0.015))) THEN
            ERRORF = 5
         ELSE IF ((MW.GT.1000).AND.(VB.GT.(0.5))) THEN
            ERRORF = 6
         ELSE IF (MW.GT.1000) THEN
            ERRORF = 7
         ELSE IF (VB.LT.(0.015)) THEN
            ERRORF = 8
         ELSE IF (VB.GT.(0.5)) THEN
            ERRORF = 9
         END IF

      END

!C*************************************************************


