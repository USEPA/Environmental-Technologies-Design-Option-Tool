!C****************************************************************
!CC
!CC                      PT1DTOW
!CC
!CC Description:  This subroutine will calculate the diameter of
!CC               the tower, given tower area.
!CC
!CC Output Variable:
!CC    DT =       Tower Diameter (m)
!CC
!CC Input Variable:
!CC    AREA =     Tower area (m^2)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C****************************************************************

      SUBROUTINE PT1DTOW(DT,AREA)
!C  ATTRIBUTES DLLEXPORT, STDCALL::PT1DTOW
!C  ATTRIBUTES ALIAS:'_PT1DTOW@8':: PT1DTOW
!C  ATTRIBUTES REFERENCE::DT,AREA

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DT,AREA

         DT = ((4*AREA)/(3.1415926))**0.5

      END

!C****************************************************************

