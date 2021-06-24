!CC********************************************************************
!CC
!CC                             PRESSCNV
!CC             CONVERT PRESSURE UNITS FROM Pa TO psi
!CC
!CC Description:  This SUBROUTINE will convert pressure from
!CC               units of Pa to units of psi.
!CC
!CC Output Variables:
!CC    PRESSENG =    Vapor Pressure (psi)
!CC
!CC Input Variables:
!CC    PRESSSI =     Vapor Pressure (Pa)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/21/94)
!CC
!CC********************************************************************

      SUBROUTINE PRESSCNV(PRESSENG,PRESSSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PRESSCNV
!MS$ ATTRIBUTES ALIAS:'_PRESSCNV'::PRESSCNV
!MS$ ATTRIBUTES REFERENCE::PRESSENG,PRESSSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION PRESSENG, PRESSSI

         PRESSENG = PRESSSI * 14.696D0 / 1.01325D+05

      END

!CC********************************************************************


