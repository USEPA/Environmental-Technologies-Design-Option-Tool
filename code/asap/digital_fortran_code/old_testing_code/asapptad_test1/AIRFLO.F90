!C***************************************************************
!CC
!CC                            AIRFLO
!CC         CALCULATION OF THE AIR FLOW RATE TO EACH TANK
!CC
!CC Description:  This subroutine will calculate the air flow
!CC               rate (to each tank for bubble aeration and to
!CC               the tower for packed tower aeration), given
!CC               air to water ratio and water flow rate
!CC
!CC Output Variable:
!CC    QA =       Air flow rate (m^3/sec)
!CC
!CC Input Variable:
!CC    VQ =       Air to water ratio (dimensionless)
!CC    QW =       Water flow rate (m^3/sec)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE AIRFLO(QA,VQ,QW)
!C  ATTRIBUTES DLLEXPORT, STDCALL::AIRFLO
!C  ATTRIBUTES ALIAS:'_AIRFLO@12':: AIRFLO
!C  ATTRIBUTES REFERENCE::QA,VQ,QW

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION QA,VQ,QW

         QA = VQ*QW

      END

!C***************************************************************

