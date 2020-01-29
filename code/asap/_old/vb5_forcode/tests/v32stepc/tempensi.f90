!C*******************************************************************
!C
!C                              TEMPENSI
!C               CONVERT TEMPERATURE UNITS FROM F TO C
!C
!C Description:  This SUBROUTINE will convert temperature from units of 
!C               F to units of C. 
!C
!C Output Variables:
!C    TEMPSI =     Temperature  (C)
!C
!C Input Variables:
!C    TEMPENG =    Temperature (F)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE TEMPENSI(TEMPSI,TEMPENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::TEMPENSI
!MS$ ATTRIBUTES ALIAS:'_TEMPENSI':: TEMPENSI
!MS$ ATTRIBUTES REFERENCE::TEMPSI,TEMPENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION TEMPENG, TEMPSI

TEMPSI = (TEMPENG - 32D0) * 5D0 / 9D0
END SUBROUTINE
 
!C*******************************************************************


       
