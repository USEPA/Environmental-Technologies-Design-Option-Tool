!CC*******************************************************************
!CC
!CC                              TEMPENSI
!CC               CONVERT TEMPERATURE UNITS FROM F TO C
!CC
!CC Description:  This SUBROUTINE will convert temperature from units of
!CC               F to units of C.
!CC
!CC Output Variables:
!CC    TEMPSI =     Temperature  (C)
!CC
!CC Input Variables:
!CC    TEMPENG =    Temperature (F)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE TEMPENSI(TEMPSI,TEMPENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::TEMPENSI
!MS$ ATTRIBUTES ALIAS:'_TEMPENSI'::TEMPENSI
!MS$ ATTRIBUTES REFERENCE::TEMPSI,TEMPENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION TEMPENG, TEMPSI
        TEMPSI = (TEMPENG - 32D0) * 5D0 / 9D0
      END
 
!CC*******************************************************************


       
