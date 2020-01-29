!C***************************************************************
!CC
!CC                    PTOTALPT
!CC
!CC Description:  This subroutine will calculate the total power,
!CC               which is equal to the sum of the brake power for
!CC               the blower and the brake power for the pump.
!CC
!CC Output Variable:
!CC    BPT =      Total brake power (kW)
!CC
!CC Input Variables:
!CC    BP =       Brake power for the blower (kW)
!CC    BPW =      Brake power for the pump (kW)
!CC
!CC History:
!CC    Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PTOTALPT(BPT,BP,BPW)        
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PTOTALPT
!MS$ ATTRIBUTES ALIAS:'_PTOTALPT':: PTOTALPT
!MS$ ATTRIBUTES REFERENCE::BPT,BP,BPW

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION BPT,BP,BPW

         BPT = BP + BPW

      END

!C***************************************************************

