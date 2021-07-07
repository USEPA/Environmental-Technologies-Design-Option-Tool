!C***************************************************************
!CC
!CC                      PT1VQMIN
!CC
!CC Description:  This subroutine will calculate the minimum air
!CC               to water ratio to achieve the treatment objective
!CC               for a compound of interest, given influent conc.,
!CC               effluent conc., and Henry's Constant
!CC
!CC Output Variable:
!CC    VQMIN =    Minimum air to water ratio (dimensionless)
!CC
!CC Input Variables:
!CC    CI =       Influent concentration (ug/L)
!CC    CE =       Effluent concentration (ug/L)
!CC    HC =       Henry's Constant (dimensionless)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PT1VQMIN(VQMIN,CI,CE,HC)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PT1VQMIN
!MS$ ATTRIBUTES ALIAS:'_PT1VQMIN@16':: PT1VQMIN
!MS$ ATTRIBUTES REFERENCE::VQMIN,CI,CE,HC

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VQMIN,CI,CE,HC 
         
         VQMIN = (CI-CE)/(HC*CI)                                 

      END

!C***************************************************************

