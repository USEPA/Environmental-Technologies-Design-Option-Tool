!C*************************************************************
!CC
!CC                    DIFLWC
!CC
!CC Description:  This subroutine will calculate liquid diffusivity
!CC               for compounds using the Wilke-Chang correlation.
!CC               It is valid in for a solute in very dilute solution
!CC               of solvent.
!CC
!CC Output Variables:
!CC    DIFL =     Liquid diffusivity value (m^2/sec) of compound
!CC    ERRORF =   Error flag
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    DIFLT =    Temperature of this calculation (C)
!CC
!CC Input Variables:
!CC    VB =       Molar volume of compound at normal boiling
!CC               point (m^3/kmol)
!CC    TEMPOP =   Temperature of the calculation (C)
!CC
!CC Variables Internal to Subroutine DIFLWC
!CC    MWT =      Molecular Weight of Solvent (water)
!CC    VL =       Viscosity of solution (kg/m/sec)
!CC    PHI =      Association parameter of solvent
!CC               = 2.26 for water
!CC    TT =       Temperature of the calculation (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (10/19/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C*************************************************************

      SUBROUTINE DIFLWC(DIFL,VB,TEMPOP,ERRORF,SRCSHT,SRCLNG,DIFLT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DIFLWC
!MS$ ATTRIBUTES ALIAS:'_DIFLWC':: DIFLWC
!MS$ ATTRIBUTES REFERENCE::DIFL,VB,TEMPOP,ERRORF,SRCSHT,SRCLNG,DIFLT

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFL,VL,VB,TEMP,MWT,TEMPOP,PHI
      INTEGER ERRORF,SRCSHT,SRCLNG

         ERRORF = 0
         SRCSHT = 12
         DIFLT = TEMPOP
         TT = DIFLT + 273.15D0
         CALL H2OVISC(VL,TEMPOP,MERR,IDUMSHT,IDUMLNG,DUMTMP)
         MWT = 18.02D0                                        
         PHI = 2.26D0
         DIFL = (117.3D-18)*((PHI*MWT)**0.5)*TT/VL/(VB**0.6)

      END

!C*************************************************************

