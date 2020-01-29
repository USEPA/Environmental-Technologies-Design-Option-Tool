!C***************************************************************
!CC
!CC                      SEFFL
!CC
!CC Description:  This subroutine will determine the effluent
!CC               concentrations from each tank for surface
!CC               aeration.  It will also calculate the removal
!CC               efficiency.
!CC
!CC Output Variables:
!CC    CEXIT =    Array of effluent concentrations from each tank (ug/L)
!CC    RECE =     Achieved removal efficiency (%)
!CC
!CC Input Variables:
!CC    CI =       Influent concentration to tank 1 (ug/L)
!CC    KLA =      Compound mass transfer coefficient (1/sec)
!CC    TAUI =     Residence time of each tank (hrs)
!CC    NTANK =    No. of tanks (in series)
!CC
!C***************************************************************

      SUBROUTINE SEFFL(CEXIT,RECE,CI,KLA,TAUI,NTANK)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::SEFFL
!MS$ ATTRIBUTES ALIAS:'_SEFFL':: SEFFL
!MS$ ATTRIBUTES REFERENCE::CEXIT,RECE,CI,KLA,TAUI,NTANK

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DIMENSION CEXIT(20)
         INTEGER NTANK
         DOUBLE PRECISION CI,KLA,TAUI,CEXIT

         DO 1900, I=1,NTANK
            CEXIT(I) = CI/((1.0D0+KLA*TAUI*3600.0D0)**I)
1900     CONTINUE
         RECE = 100.0D0 * (CI-CEXIT(NTANK)) / CI

      END

!C***************************************************************

