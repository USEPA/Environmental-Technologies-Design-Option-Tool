!C**************************************************************
!CC
!CC                     AIRDENS
!CC
!CC Description:  This subroutine will estimate the air
!CC               density using an equation developed assuming
!CC               air is an ideal gas.
!CC
!CC Output Variables:
!CC    DG =       Air Density (kg/m^3)
!CC    ERRORF =
!CC    SRCSHT =
!CC    SRCLNG =
!CC    DGTEMP =
!CC
!CC Input Variables:
!CC    TEMPOP =   Temperature of the calculation (C)
!CC    PRESOP =   Operating pressure (N/m2)
!CC
!CC Variables Internal to Subroutine AIRDENS:
!CC    MWAVG =    Average molecular weight of air
!CC    R =        Universal Gas Constant
!CC    TEMP =     Temperature of the calculation (K)
!CC    PRES =     Pressure of the calculation (atm)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C**************************************************************

      SUBROUTINE AIRDENS(DG,TEMPOP,PRESOP,ERRORF,SRCSHT,SRCLNG,DGTEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AIRDENS
!MS$ ATTRIBUTES ALIAS:'_AIRDENS':: AIRDENS
!MS$ ATTRIBUTES REFERENCE::DG,TEMPOP,PRESOP,ERRORF,SRCSHT,SRCLNG,DGTEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DG,TEMP,R,MWAVG,PRES,TEMPOP 
      INTEGER ERRORF,SRCSHT,SRCLNG
      
         ERRORF = 0
         SRCSHT = 16
         DGTEMP = TEMPOP
         TEMP = TEMPOP + 273.15D0
         PRES = PRESOP / 1.01325D+05
         MWAVG = 28.95D0
         R = 0.08205D0
         DG = ((MWAVG)*(PRES))/((R)*(TEMP))

      END

!C**************************************************************


