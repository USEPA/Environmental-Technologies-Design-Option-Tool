!CC*******************************************************************
!CC
!CC                              NBPENSI
!CC             CONVERT NORMAL BOILING POINT UNITS FROM F TO C
!CC
!CC Description:  This SUBROUTINE will convert normal boiling point from
!CC               units of F to units of C.
!CC
!CC Output Variables:
!CC    NBPSI =     Normal Boiling Point (C)
!CC
!CC Input Variables:
!CC    NBPENG =    Normal Boiling Point (F)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE NBPENSI(NBPSI,NBPENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::NBPENSI
!MS$ ATTRIBUTES ALIAS:'_NBPENSI'::NBPENSI
!MS$ ATTRIBUTES REFERENCE::NBPSI,NBPENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION NBPENG, NBPSI
        NBPSI = (NBPENG - 32D0) * 5D0 / 9D0
      END
 
!CC*******************************************************************


       
