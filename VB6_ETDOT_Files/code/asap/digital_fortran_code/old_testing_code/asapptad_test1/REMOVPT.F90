!C***************************************************************
!CC
!CC                       REMOVPT
!CC
!CC Description:  This subroutine will calculate the
!CC               the removal efficiency for a given
!CC               compound specified earlier.
!CC
!CC Output Variables:
!CC    REMOV =    Removal Efficiency (%)
!CC
!CC Input Variables:
!CC    CI =       Influent concentration (ug/L)
!CC    CE =       Effluent concentration (ug/L)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE REMOVPT(REMOV,CI,CE)
!C  ATTRIBUTES DLLEXPORT, STDCALL::REMOVPT
!C  ATTRIBUTES ALIAS:'_REMOVPT':: REMOVPT
!C  ATTRIBUTES REFERENCE::REMOV,CI,CE

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION REMOV,CI,CE

         REMOV = ((CI-CE)/CI)*100.0D0

      END

!C***************************************************************

