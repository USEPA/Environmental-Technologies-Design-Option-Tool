!CC***************************************************************************
!CC
!CC                                 HENRY
!CC                       CALCULATE HENRY'S CONSTANT
!CC
!CC Output Variables:
!CC    HLC =      Henry's constant (-)
!CC
!CC Input Variables:
!CC    GAMMA =    Activity coefficient (-)
!CC    PVAP =     Vapor pressure (mm Hg)
!CC    TEMP =     Temperature of calculation (K)
!CC    PSAT =     Saturation pressure (atm)
!CC
!CC Variables Internal to Subroutine HENRY:
!CC    HATM =      Henry's constant (atm)
!CC
!CC Authors:  M. Miller, T. Rogers, D. Hokanson
!CC
!CC **************************************************************************

      SUBROUTINE HENRY (TEMP,HLC,GAMMA,PVAP,NG,PSAT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HENRY
!MS$ ATTRIBUTES ALIAS:'_HENRY@24':: HENRY
!MS$ ATTRIBUTES REFERENCE::TEMP,HLC,GAMMA,PVAP,NG,PSAT

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      PSAT = PVAP/760.0D0
      HATM = GAMMA*PSAT
      HLC = HATM*(18.015D0/1000.0D0)/(0.082054D0*TEMP)

      END


