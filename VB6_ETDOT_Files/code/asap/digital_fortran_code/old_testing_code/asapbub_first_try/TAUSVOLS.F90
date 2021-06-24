!C****************************************************************
!CC
!CC                     TAUSVOLS
!CC
!CC Description:  This subroutine will calculate the fluid
!CC               residence time for each tank, the total
!CC               fluid residence time, the individual tank
!CC               volume and the total volume of all tanks
!CC               for bubble and surface aeration.  Given any
!CC               one of the 4 parameters, the other three will be
!CC               calculated.
!CC
!CC Output Variables:
!CC    TAU =      Fluid residence time in each tank (hrs)
!CC    TAUN =     Total fluid residence time (hrs)
!CC    VTANK =    Volume of each tank (m3)
!CC    VTOT =     Total volume of all tanks (m3)
!CC
!CC Input Variables:
!CC    NTANK =    No. of tanks in series
!CC    QW =       Water Flow Rate (m3/sec)
!CC    CODE =     Code telling which variable is user-specified
!CC               and which variables must be calculated:
!CC                  Code = 1 --> Specified TAU
!CC                               Calculate TAUN, VTANK, VTOT
!CC                  Code = 2 --> Specified TAUN
!CC                               Calculate TAU, VTANK, VTOT
!CC                  Code = 3 --> Specified VTANK
!CC                               Calculate TAU, TAUN, VTOT
!CC                  Code = 4 --> Specified VTOT
!CC                               Calculate TAU, TAUN, VTANK
!CC
!C****************************************************************

      SUBROUTINE TAUSVOLS(TAUN,NTANK,TAU,VTANK,VTOT,QW,CODE)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::TAUSVOLS
!MS$ ATTRIBUTES ALIAS:'_TAUSVOLS':: TAUSVOLS
!MS$ ATTRIBUTES REFERENCE::TAUN,NTANK,TAU,VTANK,VTOT,QW,CODE

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         INTEGER NTANK, CODE
         DOUBLE PRECISION TAUN,TAU,VTANK,VTOT,QW

         IF (CODE.EQ.1) THEN
            TAUN = DBLE(NTANK)*TAU
            VTANK = TAU*60.0D0*60.0D0*QW
            VTOT = VTANK*DBLE(NTANK)
         ELSE IF (CODE.EQ.2) THEN
            TAU = TAUN / DBLE(NTANK)
            VTANK = TAU*60.0D0*60.0D0*QW
            VTOT = VTANK*DBLE(NTANK)       
         ELSE IF (CODE.EQ.3) THEN
            TAU = VTANK / 60.0D0 / 60.0D0 / QW
            TAUN = DBLE(NTANK)*TAU
            VTOT = VTANK*DBLE(NTANK)
         ELSE IF (CODE.EQ.4) THEN
            VTANK = VTOT / DBLE(NTANK)
            TAU = VTANK / 60.0D0 / 60.0D0 / QW
            TAUN = DBLE(NTANK)*TAU
         END IF

      END

!C****************************************************************

