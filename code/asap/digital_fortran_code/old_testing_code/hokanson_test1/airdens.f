

!C**************************************************************
!C
!C                     AIRDENS
!C
!C Description:  This subroutine will estimate the air
!C               density using an equation developed assuming
!C               air is an ideal gas.
!C
!C Output Variable:
!C    DG =       Air Density (kg/m^3)
!C
!C Input Variables:
!C    TEMP =     Temperature in Deg K
!C    PRES =     Operating pressure (atm)
!C
!C Variables Internal to Subroutine AIRDENS:
!C    MWAVG =    Average molecular weight of air
!C    R =        Universal Gas Constant
!C
!C History:  Subroutine written by David R. Hokanson (9/30/93)
!C
!C**************************************************************

      SUBROUTINE AIRDENS(DG,TEMP,PRES)
!DEC$ ATTRIBUTES DLLEXPORT :: AIRDENS
!DEC$ ATTRIBUTES ALIAS : "_AIRDENS" :: AIRDENS

!cccccccc MS ATTRIBUTES DLLEXPORT, STDCALL::AIRDENS
!cccccccc MS ATTRIBUTES ALIAS:'_AIRDENS':: AIRDENS
!cccccccc MS ATTRIBUTES REFERENCE::DG,TEMP,PRES

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DG,TEMP,R,MWAVG,PRES 
      
         MWAVG = 28.95D0
         R = 0.08205D0
         DG = ((MWAVG)*(PRES))/((R)*(TEMP))

      END

!C**************************************************************

