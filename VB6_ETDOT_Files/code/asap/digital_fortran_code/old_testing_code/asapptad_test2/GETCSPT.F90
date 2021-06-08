!C***************************************************************
!CC
!CC                  GETCSPT
!CC
!CC Description:  This subroutine will calculate the concentration
!CC               at the air-water interface
!CC
!CC Output Variable:
!CC    CS =       Conc. at air-water interface (ug/L)
!CC
!CC Input Variables:
!CC    VQ =       Air to water ratio (dimensionless)
!CC    HC =       Henry's constant (dimensionless)
!CC    CI =       Influent concentration (ug/L)
!CC    CE =       Treatment objective (ug/L)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE GETCSPT(CS,VQ,HC,CI,CE)
!C...IBUTES ALIAS:'_GETCSPT@20':: GETCSPT

!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETCSPT
!MS$ ATTRIBUTES ALIAS:'_GETCSPT@20':: GETCSPT
!MS$ ATTRIBUTES REFERENCE::CS,VQ,HC,CI,CE

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION CS,VQ,HC,CI,CE

         CS = (1.0D0/(VQ*HC)) * (CI-CE)

      END

!C***************************************************************

