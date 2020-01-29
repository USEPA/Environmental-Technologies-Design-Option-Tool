!C***************************************************************
!CC
!CC                         GETPHIB
!CC
!CC Description:  This subroutine will calculate the Stanton no.,
!CC               PHI, for bubble aeration.
!CC
!CC Output Variables:
!CC    PHI =      Stanton No. (>3 --> Mass transfer not controlling)
!CC
!CC Input Variables:
!CC    KLA =      Compound mass transfer coefficient (1/sec)
!CC    VTANK =    Volume of each tank (m^3)
!CC    HC =       Henry's constant of compound (dimensionless)
!CC    QA =       Air flow rate to each tank (m^3/sec)
!CC
!C***************************************************************

      SUBROUTINE GETPHIB(PHI,KLA,VTANK,HC,QA)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETPHIB
!MS$ ATTRIBUTES ALIAS:'_GETPHIB':: GETPHIB
!MS$ ATTRIBUTES REFERENCE::PHI,KLA,VTANK,HC,QA
         
         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION PHI,KLA,VTANK,HC,QA

         PHI = KLA*VTANK/HC/QA 

      END

!C***************************************************************

