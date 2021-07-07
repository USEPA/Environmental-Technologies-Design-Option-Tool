!CC*********************************************************************
!CC
!CC                               REGRESS
!CC                      HENRY'S CONSTANT REGRESSION
!CC
!CC Output Variable:
!CC    HR =        Henry's constant regression value (atm)
!CC
!CC Input Variables:
!CC    TT =        Temperature at which regression value desired (K)
!CC    HC =        Array of Henry's constant values in database (atm)
!CC    TMP1 =      Array of Henry's constant temperature pts in database
!CC    NUMDBHCS =  Number of Henry's constant data points in database
!CC
!CC Variables Internal to Subroutine REGRESS:
!CC    SUMXY =
!CC    SUMY =
!CC    SUMX =
!CC    SUMX2 =
!CC    CEPT =
!CC    SLOPE =
!CC    TMPHR =
!CC
!CC Authors:  M. Miller and T. Rogers (4/4/94)
!CC
!CC*********************************************************************

      SUBROUTINE REGRESS (TT,HC,TMP1,HR,NUMDBHCS)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::REGRESS
!MS$ ATTRIBUTES ALIAS:'_REGRESS@20':: REGRESS
!MS$ ATTRIBUTES REFERENCE::TT,HC,TMP1,HR,NUMDBHCS

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      PARAMETER (ND=20)

      DIMENSION HC(ND),TMP1(ND)

!CC    -- THIS REGRESSION WAS DONE BY PLOTTING LN OF HENRY'S DATA VS 1/T
!CC    -- INITIALIZE VARIABLES

      SUMXY = 0.0D0
      SUMY = 0.0D0
      SUMX = 0.0D0
      SUMX2 = 0.0D0
      CEPT = 0.0D0
      SLOPE = 0.0D0
      TMPHR = 0.0D0

!CC    -- FIND SUMS TO DO REGRESSION --

      DO 10 I=1, NUMDBHCS

           SUMXY = SUMXY + DLOG(HC(I))*(1.0D0/(TMP1(I)+273.15D0))
           SUMY = SUMY + DLOG(HC(I))
           SUMX = SUMX + 1.0D0/(TMP1(I)+273.15D0)
           SUMX2 = SUMX2 + (1.0D0/(TMP1(I)+273.15D0))**2
          
10    CONTINUE

!CC    -- CALCULATE MEANS, SLOPE, AND INTERCEPT --

      RMEANX = SUMX/NUMDBHCS
      RMEANY = SUMY/NUMDBHCS

      SLOPE = ((NUMDBHCS*SUMXY)-(SUMX*SUMY))/((NUMDBHCS*SUMX2)-SUMX**2)
      CEPT =  RMEANY - SLOPE*RMEANX

!CC    -- FIND NEW HENRY'S CONSTANT AT OPERATING T --

      TMPHR = SLOPE*(1.0D0/TT) + CEPT
      HR = DEXP(TMPHR)

      END


