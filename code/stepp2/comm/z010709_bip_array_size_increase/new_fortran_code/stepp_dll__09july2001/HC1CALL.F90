!CC*********************************************************************
!CC
!CC                      HC1CALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to get the UNIFAC value of Henry's
!CC   Constant at the operating T.
!CC
!CC Output Variables:
!CC    VAL =       Henry's constant value (-)
!CC    SRCSHT =    Source of the value (Short Version)
!CC    SRCLNG =    Source of the value (Long Version)
!CC    ERRORF =    Error Flag
!CC    TEMPUN =    Temperature of the value (C)
!CC
!CC Input Variables:
!CC    TEMPOP =    Operating temperature (C)
!CC    GAMMA =     Activity Coefficient (-)
!CC    PVAP =      Vapor pressure (N/m2)
!CC
!CC Variables Internal to Subroutine HC1CALL:
!CC    VP =        Vapor pressure (mm Hg)
!CC    GAMMA =     Activity coefficient (-)
!CC    TT =        Operating temperature (K)
!CC    HLC =       Henry's constant value returned from subroutine HENRY
!CC    PSAT =      Saturation pressure (atm)
!CC
!CC Author:  D. Hokanson (4/4/94)
!CC
!CC*********************************************************************
       
      SUBROUTINE HC1CALL(VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMPOP,GAMMA,PVAP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HC1CALL
!MS$ ATTRIBUTES ALIAS:'_HC1CALL':: HC1CALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMPOP,GAMMA,PVAP

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)

!CC--------Modified David R. Hokanson 7/9/01 for STEPP2
!CC--------   Increased dimensioning for new binary interaction parameter databases      
!CC         PARAMETER(MA=53,NA=96,NC=2,ND=10)
         PARAMETER(MA=58,NA=116,NC=2,ND=10)
!CC--------End Modified David R. Hokanson 7/9/01 for STEPP2
               
         COMMON /INIT/ XX(10), NG, NDIF
         
         DOUBLE PRECISION VAL,TEMPUN,TEMPOP,GAMMA,PVAP,AC,VP
         INTEGER SRCSHT,SRCLNG,ERRORF

         ERRORF = 0
         SRCSHT = 7
         TEMPUN = TEMPOP

         TT = TEMPUN + 273.15D0
         AC = GAMMA * 1.0D0
         VP = PVAP * 760.0D0 / 1.01325D+05
         CALL HENRY(TT,HLC,AC,VP,NG,PSAT)
         VAL = HLC
         IF (PSAT.GT.(1.0D0)) THEN 
           ERRORF = 10
         ENDIF

         END

!CC************************************************************


