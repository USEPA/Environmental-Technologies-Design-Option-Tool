!C***************************************************************
!CC
!CC                         REMOVBUB
!CC
!CC Description:  This subroutine will calculate the overall
!CC               overall removal efficiency for the liquid
!CC               phase.
!CC
!CC Output Variables:
!CC    RECE =     Actual liquid phase removal efficiency
!CC
!CC Input Variables:
!CC    CEFFL =    Liquid phase effluent conc. from last tank (ug/L)
!CC    CI =       Liquid phase influent conc. (ug/L)
!CC
!C***************************************************************

      SUBROUTINE REMOVBUB(RECE,CI,CEFFL)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::REMOVBUB
!MS$ ATTRIBUTES ALIAS:'_REMOVBUB':: REMOVBUB
!MS$ ATTRIBUTES REFERENCE::RECE,CI,CEFFL
         
         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION RECE,CI,CEFFL

         RECE = (CI-CEFFL)/CI * 100.0D0

      END

!C***************************************************************

