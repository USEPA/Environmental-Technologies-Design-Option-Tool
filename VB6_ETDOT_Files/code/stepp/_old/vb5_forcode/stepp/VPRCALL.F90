!CC************************************************************
!CC
!CC                      VPRCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to get vapor pressure values
!CC
!CC Output Variables:
!CC    VAL =        Vapor Pressure value (N/m^2)
!CC    SRCSHT =     Source of the value (Short Version)
!CC    SRCLNG =     Source of the value (Long Version)
!CC    ERRORF =     Error Flag
!CC    TEMPDB =     Temperature of the database Value from DIPPR801
!CC                 or Yaws (C)
!CC    TMIN =       Minimum valid database temperature for DIPPR801
!CC                 or Yaws (C)
!CC    TMAX =       Maximum valid database temperature for DIPPR801
!CC                 or Yaws (C)
!CC
!CC Input Variables:
!CC    NEQN =       Number of the Vapor Pressure equation (database)
!CC    ANTA =       Database Vapor Pressure correlation coefficient A
!CC    ANTB =       Database Vapor Pressure correlation coefficient B
!CC    ANTC =       Database Vapor Pressure correlation coefficient C
!CC    ANTD =       Database Vapor Pressure correlation coefficient D
!CC    ANTE =       Database Vapor Pressure correlation coefficient E
!CC    VPSF =       Database Vapor Pressure (from Superfund)
!CC    TEMPSF =     Database Vapor Pressure (from Superfund)
!CC    TEMPOP =     Operating temperature (C)
!CC
!CC Variables Internal to Subroutine VPRCALL:
!CC    PVAP =       Vapor pressure returned from subroutine VAPORP
!CC                 (units vary) for Database DIPPR801 or Yaws
!CC    TT =         Operating temperature (K)
!CC
!CC Author:  D. Hokanson (4/3/94)
!CC
!CC************************************************************
       
      SUBROUTINE VPRCALL(VAL,SRCSHT,SRCLNG,ERRORF,NEQN,TEMPDB,TMIN,TMAX,ANTA,ANTB,ANTC,ANTD,ANTE,VPSF,TEMPSF,TEMPOP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VPRCALL
!MS$ ATTRIBUTES ALIAS:'_VPRCALL':: VPRCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,NEQN,TEMPDB,TMIN,TMAX,ANTA,ANTB,ANTC,ANTD,ANTE,VPSF,TEMPSF,TEMPOP

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
         DOUBLE PRECISION VAL,TEMPDB,TMIN,TMAX,ANTA,ANTB,ANTC,ANTD,ANTE,VPSF,TEMPSF,TEMPOP
         INTEGER SRCSHT,SRCLNG,ERRORF,NEQN

         ERRORF = 0
         TT = TEMPOP + 273.15D0
         IF ((ANTA.LT.0).AND.(VPSF.LT.0)) THEN
            ERRORF = -1
         ELSE IF (ANTA.LT.0) THEN 
            VAL = VPSF * 1.01325D+05 / 760.0D0
            ERRORF = 2
            TEMPDB = TEMPSF
         ELSE IF ((NEQN.NE.101).AND.(NEQN.NE.-1)) THEN
            ERRORF = -8
         ELSE
            NOVPT = 0
            CALL VAPORP(PVAP,TT,ANTA,ANTB,ANTC,ANTD,ANTE,NOVPT,NEQN,TMIN,TMAX,SRCSHT)

            IF (NOVPT.EQ.-1) THEN
               ERRORF = 1
            END IF
            
            IF (SRCSHT.EQ.4) THEN
!CC              *** DIPPR801
               VAL = PVAP
            ELSE
!CC              *** YAWS
               VAL = PVAP * 1.01325D+05 / 760.0D0
            END IF   
               
            TEMPDB = TEMPOP
         END IF
           
        RETURN
      END

!CC************************************************************

