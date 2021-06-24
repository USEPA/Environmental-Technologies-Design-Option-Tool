!C***************************************************************
!CC
!CC                       SURFEFF
!CC
!CC Description:  This subroutine will calculate the removal
!CC               efficiency desired for surface aeration.
!CC
!CC Output Variable:
!CC    RMOVAL =   Desired removal efficiency for surface aeration (%)
!CC
!CC Input Variables:
!CC    CI =       Influent concentration (ug/L)
!CC    CE =       Treatment objective (ug/L)
!CC
!C***************************************************************

      SUBROUTINE SURFEFF(RMOVAL,CI,CE)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::SURFEFF
!MS$ ATTRIBUTES ALIAS:'_SURFEFF':: SURFEFF
!MS$ ATTRIBUTES REFERENCE::RMOVAL,CI,CE

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION RMOVAL,CI,CE

         RMOVAL = (CI-CE)*100.0D0/CI

      END

!C***************************************************************

