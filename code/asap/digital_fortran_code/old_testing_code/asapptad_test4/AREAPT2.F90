!C***************************************************************
!CC
!CC                       AREAPT2
!CC
!CC Description:  This subroutine will calculate tower area,
!CC               given tower diameter.
!CC
!CC Output Variables:
!CC    AREA =     Tower area (m^2)
!CC
!CC Input Variables:
!CC    DT =       Tower diameter (m)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE AREAPT2(AREA,DT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AREAPT2
!MS$ ATTRIBUTES ALIAS:'_AREAPT2':: AREAPT2
!MS$ ATTRIBUTES REFERENCE::AREA,DT

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION AREA,DT

         AREA = 3.1415926D0 * DT * DT / 4.0D0

      END

!C***************************************************************

