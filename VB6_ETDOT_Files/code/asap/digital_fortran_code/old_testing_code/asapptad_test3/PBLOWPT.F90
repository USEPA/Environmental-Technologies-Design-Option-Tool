!C*************************************************************
!CC
!CC                         PBLOWPT
!CC
!CC Description:  This subroutine calculates the brake power
!CC               for the blowers used to supply air to the
!CC               packed tower.
!CC
!CC Output Variable:
!CC    BP =       Brake power for the blower (kW)
!CC
!CC Input Variables:
!CC    QA =       Air flow rate (m^3/sec)
!CC    AREA =     Tower area (m^2)
!CC    PRES =     Operating pressure (atm)
!CC    PRESD =    Gas pressure drop (N/m^2/m)
!CC    HLL =      Tower length (m)
!CC    DG =       Density of air (kg/m^3)
!CC    T1 =       Inlet air temperature (Deg C)
!CC    EFF =      Blower efficiency (%)
!CC
!CC Variables Internal to Subroutine PBLOWPT
!CC    VGAS =     Gas volumetric loading (m^3/m^2/sec)
!CC    PRESE =    Pressure drop through demister, packing support
!CC               plate, duct work, and tower inlet and outlet (N/m^2)
!CC    RG =       Universal gas constant for air (J/kg air/K)
!CC    NN =       Constant (=0.283 for air)
!CC    P1 =       Outlet air pressure (N/m^2)
!CC    P2 =       Inlet air pressure (N/m^2)
!CC    QMG =      Mass flow rate of air (kg/sec)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C*************************************************************

      SUBROUTINE PBLOWPT(BP,QA,AREA,PRES,PRESD,HLL,DG,T1,EFF)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PBLOWPT
!MS$ ATTRIBUTES ALIAS:'_PBLOWPT':: PBLOWPT
!MS$ ATTRIBUTES REFERENCE::BP,QA,AREA,PRES,PRESD,HLL,DG,T1,EFF

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)  
      DOUBLE PRECISION BP,QA,AREA,PRES,PRESD,HLL,DG,T1,EFF,VGAS,PRESE,RG,NN,P1,P2,QMG

      T1 = T1+273.0D0
      EFF = EFF / 100.0D0
      VGAS = QA/AREA
      PRESE = 275.0D0 * (VGAS**2)
      RG = 286.7D0
      NN = 0.283D0
      P1 = PRES * (101330.0D0)
      P2 = (PRESD*HLL) + P1 + PRESE
      QMG = QA*DG
      BP = ((QMG*RG*T1)/(1000.0D0*NN*EFF))*((P2/P1)**NN -1)
      T1 = T1 - 273.0D0
      EFF = EFF * 100.0D0

      END

!C*************************************************************

