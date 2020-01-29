!C********************************************************************
!C
!C                             NBPCONV
!C          CONVERT NORMAL BOILING POINT UNITS FROM C TO F  
!C
!C Description:  This SUBROUTINE will convert normal boiling point
!C               from units of C to units of F.
!C
!C Output Variables:
!C    NBPENG =    Normal Boiling Point (F)
!C
!C Input Variables:
!C    NBPSI =     Normal Boiling Point (C)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE NBPCONV(NBPENG,NBPSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::NBPCONV
!MS$ ATTRIBUTES ALIAS:'_NBPCONV':: NBPCONV
!MS$ ATTRIBUTES REFERENCE::NBPENG,NBPSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION NBPENG, NBPSI

NBPENG = NBPSI * (9.0D0/5.0D0) + 32.0D0         

END SUBROUTINE

!C********************************************************************
