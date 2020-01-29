!CC*********************************************************************
!CC
!CC                                 HCDBCONV
!CC      CONVERT DATABASE HENRY'S CONSTANTS INTO DIMENSIONLESS UNITS
!CC
!CC Output Variable:
!CC    HCDB =     Henry's constant database values array (-)
!CC
!CC Input Variables:
!CC    HCDB =     Henry's constant database values array
!CC                (units = atm if Yaws or RTI, atm-m3/mol if Superfund)
!CC    HCDBTMP =  Array of Henry's constant temperatures
!CC    NUMDBHCS = Number of database Henry's constants
!CC    SRCSHT =   Source of the Henry's constant values
!CC
!CC Variables Internal to Subroutine HCDBCONV
!CC    TEMP =     Temperature of current conversion (K)
!CC
!CC Author:  D. Hokanson (4/4/94)
!CC
!CC*********************************************************************

      SUBROUTINE HCDBCONV(HCDB,HCDBTMP,NUMDBHCS,SRCSHT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HCDBCONV
!MS$ ATTRIBUTES ALIAS:'_HCDBCONV':: HCDBCONV
!MS$ ATTRIBUTES REFERENCE::HCDB,HCDBTMP,NUMDBHCS,SRCSHT

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      PARAMETER (NUMHCS=20)
      DIMENSION HCDB(NUMHCS),HCDBTMP(NUMHCS)
      INTEGER SRCSHT,NUMDBHCS

         IF (SRCSHT.EQ.2) THEN
!CC           *** SUPERFUND (Convert atm-m3/mol to dimensionless)
            DO 10, I = 1,NUMDBHCS
               TEMP = HCDBTMP(I) + 273.15D0
               HCDB(I) = HCDB(I) * 1000.0D0 / 0.082054D0 / TEMP
 10         CONTINUE
         ELSE
!CC           *** YAWS OR RTI (Convert atm to dimensionless)
            DO 20, I = 1,NUMDBHCS
               TEMP = HCDBTMP(I) + 273.15D0
               HCDB(I) = HCDB(I)*(18.015D0/1000.0D0)/0.082054D0/TEMP
 20         CONTINUE
         END IF

      END


