!CC***************************************************************
!CC
!CC                        PT1AREA
!CC
!CC Description:  This subroutine will calculate tower area for the
!CC               design phase of the tower.
!CC
!CC Output Variable:
!CC    AREA =     Tower area (m^2)
!CC
!CC Input Variables:
!CC    QW =       Water flow rate (m^3/sec)
!CC    DL =       Water density (kg/m^3)
!CC    ML =       Water mass loading (kg/m^2/sec)
!CC
!CC Variable Internal to subroutine AREAPT1:
!CC    QWM =      Water mass flow rate (kg/sec)
!CC
!CC History:  Program written by David R. Hokanson (9/30/93)
!CC
!C****************************************************************

      SUBROUTINE PT1AREA(AREA,QW,DL,ML)
!C  ATTRIBUTES DLLEXPORT, STDCALL::PT1AREA
!C  ATTRIBUTES ALIAS:'_PT1AREA@16':: PT1AREA
!C  ATTRIBUTES REFERENCE::AREA,QW,DL,ML

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION AREA,QW,DL,ML,QWM

         QWM = QW*DL 
         AREA = QWM/ML

      END

!C****************************************************************

