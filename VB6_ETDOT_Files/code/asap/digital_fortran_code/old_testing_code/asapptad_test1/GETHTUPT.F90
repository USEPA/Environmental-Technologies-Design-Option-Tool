!C***************************************************************
!CC
!CC                  GETHTUPT
!CC
!CC Description:  This subroutine will calculate the height of
!CC               a transfer unit for a packed tower design.
!CC
!CC Output Variable:
!CC    HTU =      Height of a transfer unit (m)
!CC
!CC Input Variables:
!CC    QW =       Water flow rate (m^3/sec)
!CC    AREA =     Tower area (m^2)
!CC    KLA =      Overall mass transfer coefficient (1/sec)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE GETHTUPT(HTU,QW,AREA,KLA)
!C  ATTRIBUTES DLLEXPORT, STDCALL::GETHTUPT
!C  ATTRIBUTES ALIAS:'_GETHTUPT@16':: GETHTUPT
!C  ATTRIBUTES REFERENCE::HTU,QW,AREA,KLA

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION HTU,QW,AREA,KLA

         HTU = QW/(AREA*KLA)

      END

!C***************************************************************

