!C********************************************************************
!C
!C                             PRESSCNV
!C             CONVERT PRESSURE UNITS FROM Pa TO psi
!C
!C Description:  This SUBROUTINE will convert pressure from 
!C               units of Pa to units of psi.
!C
!C Output Variables:
!C    PRESSENG =    Vapor Pressure (psi)
!C
!C Input Variables:
!C    PRESSSI =     Vapor Pressure (Pa)
!C
!C History:
!C    Function written by D. Hokanson (6/21/94)
!C
!C********************************************************************

SUBROUTINE PRESSCNV(PRESSENG,PRESSSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PRESSCNV
!MS$ ATTRIBUTES ALIAS:'_PRESSCNV':: PRESSCNV
!MS$ ATTRIBUTES REFERENCE::PRESSENG,PRESSSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION PRESSENG, PRESSSI

PRESSENG = PRESSSI * 14.696D0 / 1.01325D+05

END SUBROUTINE

!C********************************************************************
