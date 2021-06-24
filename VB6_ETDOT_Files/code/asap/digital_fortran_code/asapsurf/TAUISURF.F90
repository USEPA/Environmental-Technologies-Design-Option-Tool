!C***************************************************************
!CC
!CC                         TAUISURF
!CC    RESIDENCE TIMES OF 1 TANK FOR SURFACE AERATION
!CC
!CC Description:  This subroutine will calculate the hydraulic
!CC               retention time for each individual tank   This is
!CC               for the case of a design (new) facility.
!CC
!CC Output Variables:
!CC    TAUI =     Hydraulic retention time of each tank (hrs)
!CC
!CC Input Variables:
!CC    CI =       Influent concentration (ug/L)
!CC    CE =       Treatment objective (ug/L)
!CC    NTANK =    No. of tanks (in series)
!CC    KLA =      Compound mass transfer coefficient (1/sec)
!CC
!CC Variable Internal to Subroutine TAUISURF
!CC    TAUN =     Residence time of all tanks (hrs)
!CC
!C***************************************************************

      SUBROUTINE TAUISURF(TAUI,CI,CE,NTANK,KLA)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::TAUISURF
!MS$ ATTRIBUTES ALIAS:'_TAUISURF':: TAUISURF
!MS$ ATTRIBUTES REFERENCE::TAUI,CI,CE,NTANK,KLA

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
         INTEGER NTANK
         DOUBLE PRECISION TAUN,TAUI,CI,CE,KLA

            TAUN = (NTANK/KLA) * ((CI/CE)**(1.0D0/NTANK)-1.0D0)
            TAUI = TAUN/NTANK
            TAUI = TAUI/3600.0D0
 
      END

!C***************************************************************

