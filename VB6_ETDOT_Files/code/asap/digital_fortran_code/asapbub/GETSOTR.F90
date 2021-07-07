!C********************************************************************
!CC
!CC                      GETSOTR
!CC
!CC Description:  Given SOTE and Qair, this subroutine calculates
!CC               SOTR.
!CC
!CC Output Variable:
!CC    SOTR =     Standardized oxygen mass transfer rate (kg/d)
!CC         =     Rate of oxygen mass transfer at zero D.O. and 20 Deg C
!CC
!CC Input Variables:
!CC    QAIR =     Air flow rate (std m^3/hr) --> 20 Deg C, 1 atm, 36% r.h.
!CC    SOTE =     Standardized oxygen transfer efficiency (%)
!CC
!CC Variable Internal to GETSOTR:
!CC    WO2 =      Rate of oxygen supply (kg/d) by the diffuser
!CC
!C********************************************************************

      SUBROUTINE GETSOTR (SOTR,SOTE,QAIR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETSOTR
!MS$ ATTRIBUTES ALIAS:'_GETSOTR':: GETSOTR
!MS$ ATTRIBUTES REFERENCE::SOTR,SOTE,QAIR

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION SOTR,SOTE,QAIR,WO2

         WO2 = QAIR/0.15D0
         SOTR = WO2*SOTE/100.0D0

      END

!C********************************************************************

