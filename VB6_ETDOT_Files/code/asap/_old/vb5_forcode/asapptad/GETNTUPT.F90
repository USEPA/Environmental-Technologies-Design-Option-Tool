!C***************************************************************
!CC
!CC                  GETNTUPT
!CC
!CC Description:  This subroutine will determine the number of
!CC               transfer units for a packed tower design.
!CC
!CC Output Variable:
!CC    NTU =      Number of transfer units
!CC
!CC Input Variables:
!CC    CS =       Conc. at the air-water interface (ug/L)
!CC    CI =       Influent concentration (ug/L)
!CC    CE =       Effluent concentration (ug/L)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE GETNTUPT(NTU,CI,CE,CS)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETNTUPT
!MS$ ATTRIBUTES ALIAS:'_GETNTUPT@16':: GETNTUPT
!MS$ ATTRIBUTES REFERENCE::NTU,CI,CE,CS

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION NTU,CI,CE,CS

         NTU = ((CI-CE)/(CI-CS-CE))*LOG((CI-CS)/CE)       

      END

!C***************************************************************

