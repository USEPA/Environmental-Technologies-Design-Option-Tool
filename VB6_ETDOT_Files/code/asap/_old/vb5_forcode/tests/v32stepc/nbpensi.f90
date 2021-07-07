!C*******************************************************************
!C
!C                              NBPENSI
!C             CONVERT NORMAL BOILING POINT UNITS FROM F TO C
!C
!C Description:  This SUBROUTINE will convert normal boiling point from      
!C               units of F to units of C. 
!C
!C Output Variables:
!C    NBPSI =     Normal Boiling Point (C)
!C
!C Input Variables:
!C    NBPENG =    Normal Boiling Point (F)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE NBPENSI(NBPSI,NBPENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::NBPENSI
!MS$ ATTRIBUTES ALIAS:'_NBPENSI':: NBPENSI
!MS$ ATTRIBUTES REFERENCE::NBPSI,NBPENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION NBPENG, NBPSI

NBPSI = (NBPENG - 32D0) * 5D0 / 9D0

END	SUBROUTINE

!C*******************************************************************


       
