!CC**************************************************************
!CC
!CC                      PT1TVOL
!CC
!CC Description:  This subroutine will calculate the volume
!CC               of the design tower.
!CC
!CC Output Variable:
!CC    TV =       Tower volume (m3)
!CC
!CC Input Variables:
!CC    AREA =     Tower area (m^2)
!CC    HLL =      Tower height (m)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PT1TVOL(TV,AREA,HLL)
!C  ATTRIBUTES DLLEXPORT, STDCALL::PT1TVOL
!C  ATTRIBUTES ALIAS:'_PT1TVOL@12':: PT1TVOL
!C  ATTRIBUTES REFERENCE::TV,AREA,HLL

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION TV,AREA,HLL

         TV = AREA * HLL                                

      END

!C***************************************************************

