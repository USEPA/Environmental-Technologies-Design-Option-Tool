!CC************************************************************
!CC
!CC                      LDDBCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to get liquid density values
!CC   from the equation given in the database
!CC
!CC Output Variables:
!CC     VAL =      Database liquid density value (kg/m3)
!CC     SRCSHT =   Source of this value (Short Version)
!CC     SRCLNG =   Source of this value (Long Version)
!CC     ERRORF =   Error flag
!CC     TMIN =     Minimum temp. at which correlation valid (C)
!CC     TMAX =     Maximum temp. at which correlation valid (C)
!CC
!CC Input Variables:
!CC     NEQN =     Equation number from database
!CC     TEMPDB =   Temperature at which this liquid density calculated
!CC     TMIN =     Minimum temp. at which correlation valid (K)
!CC     TMAX =     Maximum temp. at which correlation valid (K)
!CC     A =        Database liquid density coefficient A
!CC     B =        Database liquid density coefficient B
!CC     C =        Database liquid density coefficient C
!CC     D =        Database liquid density coefficient D
!CC     FWT =      Molecular weight (kg/kmol)
!CC     TEMPOP =   Operating temperature (C)
!CC
!CC Variables Internal to Subroutine LDDBCALL:
!CC    TT =        Operating temperature (K)
!CC    DBDEN =     Database liquid density (kg/m3)
!CC    NTRGE =     Error flag telling whether temp. in range
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!CC************************************************************
       
      SUBROUTINE LDDBCALL(VAL,SRCSHT,SRCLNG,ERRORF,NEQN,TEMPDB,TMIN,TMAX,A,B,C,D,FWT,TEMPOP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDDBCALL
!MS$ ATTRIBUTES ALIAS:'_LDDBCALL':: LDDBCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,NEQN,TEMPDB,TMIN,TMAX,A,B,C,D,FWT,TEMPOP

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
         DOUBLE PRECISION VAL,TEMPDB,TMIN,TMAX,A,B,C,D,FWT,TEMPOP
         INTEGER SRCSHT,SRCLNG,ERRORF,NEQN

         ERRORF = 0
         
         IF (NEQN.EQ.105) THEN
            TEMPDB = TEMPOP
            TT = TEMPDB + 273.15D0  
            NTRGE = 0
            CALL DBDENS(DBDEN,A,B,C,D,TT,TMIN,TMAX,FWT,NTRGE)
            VAL = DBDEN
            IF (NTRGE.EQ.-1) THEN
               ERRORF = 3
            END IF
         ELSE IF (NEQN.LT.0) THEN
            ERRORF = -12
         ELSE 
            ERRORF = -11
         END IF   
         
      END

!CC************************************************************

