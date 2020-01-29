!C********************************************************************
!CC
!CC                      GETSOTE
!CC
!CC Description:  Given SOTR and Qair, this subroutine calculates
!CC               SOTE.
!CC
!CC Output Variable:
!CC    SOTE =     Standardized oxygen transfer efficiency (%)C
!CC
!CC Input Variables:
!CC    QAIR =     Air flow rate (std m^3/hr) --> 20 Deg C, 1 atm, 36% r.h.
!CC
!CC    SOTR =     Standardized oxygen mass transfer rate (kg/d)
!CC         =     Rate of oxygen mass transfer at zero D.O. and 20 Deg C
!CC
!CC Variable Internal to GETSOTE:
!CC    WO2 =      Rate of oxygen supply (kg/d) by the diffuser
!CC
!C********************************************************************

      SUBROUTINE GETSOTE (SOTE,SOTR,QAIR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETSOTE
!MS$ ATTRIBUTES ALIAS:'_GETSOTE':: GETSOTE
!MS$ ATTRIBUTES REFERENCE::SOTE,SOTR,QAIR

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION SOTE,SOTR,QAIR,WO2

         WO2 = QAIR/0.15D0
         SOTE = SOTR*100.0D0/WO2

      END

!C********************************************************************

