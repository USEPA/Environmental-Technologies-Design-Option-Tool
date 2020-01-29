!C***************************************************************
!CC
!CC                        VQMINBUB
!CC      MINIMUM AIR TO WATER RATIO FOR BUBBLE AERATION
!Cc
!CC Description:   This subroutine calculates the minimum air
!CC                to water ratio for bubble aeration.  This
!CC                minimum air to water ratio is the minimum
!CC                air to water ratio to achieve the desired
!CC                removal efficiency for NTANK tanks in series.
!CC
!CC Output Variables:
!CC    VQMIN =     Minimum air to water ratio (dimensionless)
!CC
!CC Input Variable:
!CC    CINFL =     Influent Concentration (ug/L)
!CC    CTO =       Treatment Objective (ug/L)
!CC    HC =        Henry's constant (dimensionless)
!CC    NTANK =     Number of Tanks
!CC
!C***************************************************************

      SUBROUTINE VQMINBUB(VQMIN,CINFL,CTO,HC,NTANK)    
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VQMINBUB
!MS$ ATTRIBUTES ALIAS:'_VQMINBUB':: VQMINBUB
!MS$ ATTRIBUTES REFERENCE::VQMIN,CINFL,CTO,HC,NTANK

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         INTEGER NTANK
         DOUBLE PRECISION VQMIN,HC,CINFL,CTO
         DOUBLE PRECISION PARAM1

         PARAM1 = (CINFL/CTO)**(1.0D0/DBLE(NTANK))
         VQMIN = (PARAM1-1.0D0)/HC                  

      END

!C***************************************************************

