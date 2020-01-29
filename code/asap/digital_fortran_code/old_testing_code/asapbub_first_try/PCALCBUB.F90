!C******************************************************************
!CC
!CC                      PCALCBUB
!CC
!CC Description:  This subroutine does the power calculation for
!CC               bubble aeration by calculating the total blower
!CC               brake power required.
!CC
!CC Output Variables:
!CC    TP =       Total blower brake power needed for all tanks (kW)
!CC    BP =       Blower brake power required for each tank (kW)
!CC
!CC Input Variables:
!CC    PRES =     Operating pressure (atm)
!CC    TAIR =     Inlet air temperature (C)
!CC    QA =       Air flow rate (m^3/sec)
!CC    EFFB =     Blower efficiency (%)
!CC    DL =       Liquid density (kg/m^3)
!CC    HEAD =     Water depth (m)
!CC    NTANK =    No. of tanks
!CC    NBLOW =    No. of blowers per tank
!CC
!CC Variables Internal to Subroutine PCALCBUB
!CC    TAIRK =    Inlet air temperature (K)
!CC    DG =       Density of air (kg/m^3)
!CC    GME =      Air mass flow rate (kg/sec)
!CC    R =        Universal gas constant for air (J/kg/K)
!CC    NN =       0.283 for air
!CC    EFF =      Blower efficiency (as decimal)
!CC    PIN =      Inlet pressure (kN/m^2)
!CC    POUT =     Outlet pressure (kN/m^2)
!CC
!C******************************************************************

      SUBROUTINE PCALCBUB(TP,BP,PRES,TAIR,QA,EFFB,DL,HEAD,NTANK,NBLOW)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PCALCBUB
!MS$ ATTRIBUTES ALIAS:'_PCALCBUB':: PCALCBUB
!MS$ ATTRIBUTES REFERENCE::TP,BP,PRES,TAIR,QA,EFFB,DL,HEAD,NTANK,NBLOW

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      INTEGER NTANK,NBLOW
      DOUBLE PRECISION TP,BP,PRES,TAIR,QA,EFFB,DL,HEAD,DG,GME,R,NN,EFF,PIN,POUT

      TAIRK = TAIR + 273.15D0
      DG = 28.95D0*PRES/0.08205D0/TAIRK
      GME = QA*DG
      R = 286.7D0
      NN = 0.283D0
      EFF = EFFB/100.0D0
      PIN = PRES*101325.0D0
      POUT = PIN+DL*HEAD*9.81D0
      BP = (GME*R*TAIRK/1000.0D0/NN/EFF)*((POUT/PIN)**NN-1)
      TP = BP * DBLE(NTANK) * DBLE(NBLOW)

      END

!C******************************************************************

