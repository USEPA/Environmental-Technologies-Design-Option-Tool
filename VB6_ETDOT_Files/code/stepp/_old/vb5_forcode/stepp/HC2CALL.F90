!CC********************************************************************
!CC
!CC                      HC2CALL
!CC    Calculate Henry's constant at operating T by regressing data
!CC    points in the database.
!CC
!CC Output Variables:
!CC    VAL =        Henry's constant regression value (dimensionless)
!CC    SRCSHT =     Source of the value (Short Version)
!CC    SRCLNG =     Source of the value (Long Version)
!CC    ERRORF =     Error flag
!CC    TEMPREG =    Temperature of regression value (C)
!CC
!CC Input Variables:
!CC    HCDB =       Array of database Henry's constant values (-)
!CC    HCDBTMP =    Array of database temperatures (C)
!CC    TEMPOP =     Operating temperature (C)
!CC    NUMDBHCS =   Number of Henry's constant data points in database
!CC
!CC Variables Internal to Subroutine HC2CALL:
!CC    TT =         Operating temperature (K)
!CC    HCDBATM =    Array of Henry's constant values (atm)
!CC
!CC Author:  D. Hokanson (4/4/94)
!CC
!CC********************************************************************

      SUBROUTINE HC2CALL(VAL,SRCSHT,SRCLNG,ERRORF,TEMPREG,HCDB,HCDBTMP,TEMPOP,NUMDBHCS)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HC2CALL
!MS$ ATTRIBUTES ALIAS:'_HC2CALL':: HC2CALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,TEMPREG,HCDB,HCDBTMP,TEMPOP,NUMDBHCS

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      PARAMETER (NUMHCS = 20)
      DIMENSION HCDB(NUMHCS), HCDBTMP(NUMHCS), HCDBATM(NUMHCS)
      INTEGER ERRORF,SRCSHT,SRCLNG

         ERRORF = 0
         SRCSHT = 17
         TEMPREG = TEMPOP
      
         IF (NUMDBHCS.LT.2) THEN
            ERRORF = -35
         ELSE
            TT = TEMPOP + 273.15D0
            DO 10, I = 1,NUMDBHCS
                HCDBATM(I) = HCDB(I)*0.082054D0*TT*1000.0D0/18.015D0
 10         CONTINUE 
            CALL REGRESS(TT,HCDBATM,HCDBTMP,VAL,NUMDBHCS)
            VAL = VAL*(18.015D0/1000.0D0)/(0.082054D0*TT)
         END IF

      END

!CC********************************************************************


