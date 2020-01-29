!C***************************************************************
!CC
!CC                       TVOLPT2
!CC
!CC Description:  This subroutine will calculate tower volume,
!CC               given tower area and tower height
!CC
!CC Output Variables:
!CC    TV =       Tower volume (m^3)
!CC
!CC Input Variables:
!CC    AREA =     Tower area (m^2)
!CC    HLL =      Tower length (m)
!CC
!CC History:  Subroutine written by:  David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE TVOLPT2(TV,AREA,HLL)
!C  ATTRIBUTES DLLEXPORT, STDCALL::TVOLPT2
!C  ATTRIBUTES ALIAS:'_TVOLPT2':: TVOLPT2
!C  ATTRIBUTES REFERENCE::TV,AREA,HLL

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION TV,AREA,HLL

         TV = AREA * HLL

      END

!C***************************************************************

