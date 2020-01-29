!C***************************************************************
!CC
!CC                    PPUMPPT
!CC
!CC Description:  This subroutine will calculate the brake power
!CC               for the pump required to bring the water from
!CC               the bottom of the tower to the top of the
!CC               tower.
!CC
!CC Output Variable:
!CC    BPW =      Brake power for the pump (kW)
!CC
!CC Input Variables:
!CC    EFFW =     Pump efficiency (%)
!CC    DL =       Density of water (kg/m^3)
!CC    QW =       Water flow rate (m^3/sec)
!CC    HLL =      Tower height (m)
!CC
!CC History:
!CC    Subroutine written by:  David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PPUMPPT(BPW,EFFW,DL,QW,HLL) 
!C  ATTRIBUTES DLLEXPORT, STDCALL::PPUMPPT
!C  ATTRIBUTES ALIAS:'_PPUMPPT':: PPUMPPT
!C  ATTRIBUTES REFERENCE::BPW,EFFW,DL,QW,HLL

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION BPW,EFFW,DL,QW,HLL

         EFFW = EFFW/100.0D0
         BPW = (DL*QW*HLL*9.8D0)/(1000.0D0*EFFW)
         EFFW = EFFW * 100.0D0

      END

!C***************************************************************

