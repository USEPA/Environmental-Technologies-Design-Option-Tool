!C***************************************************************
!CC
!CC                      PCALCSUR
!CC          POWER CALCULATION FOR SURFACE AERATION
!CC
!CC Description:  This subroutine will perform the power calculation
!CC               for surface aeration.
!CC
!CC Output Variables:
!CC    PTOT =     Total power required (kW)
!CC    PTANK =    Power required for each tank (kW)
!CC
!CC Input Variables:
!CC    POVERV =   Power Input / Unit Volume (W/m^3)
!CC    VTOT =     Total fluid volume in all tanks (m^3)
!CC    NTANK =    No. of Tanks
!CC    EFFM =     Aerator motor efficiency (%)
!CC
!C***************************************************************

      SUBROUTINE PCALCSUR(PTOT,PTANK,POVERV,VTOT,NTANK,EFFM)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PCALCSUR
!MS$ ATTRIBUTES ALIAS:'_PCALCSUR':: PCALCSUR
!MS$ ATTRIBUTES REFERENCE::PTOT,PTANK,POVERV,VTOT,NTANK,EFFM

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         INTEGER NTANK
         DOUBLE PRECISION PTOT,PTANK,POVERV,VTOT,EFFM

         PTOT = POVERV*VTOT/(EFFM/100.0D0)/1000.0D0
         PTANK = PTOT/NTANK

      END

!C***************************************************************

