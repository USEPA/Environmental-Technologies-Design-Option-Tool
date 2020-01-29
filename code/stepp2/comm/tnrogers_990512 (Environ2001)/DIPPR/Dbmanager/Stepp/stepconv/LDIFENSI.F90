!CC*******************************************************************
!CC
!CC                                LDIFENSI
!CC          CONVERT LIQUID DIFFUSIVITY FROM Ft2/sec to m2/sec
!CC
!CC
!CC Description:  This SUBROUTINE will convert liquid diffusivity
!CC               from units of Ft2/sec to m2/sec
!CC
!CC Output Variables:
!CC    LDSI =     Liquid Diffusivity (m2/sec)
!CC
!CC Input Variables:
!CC    LDENG =    Liquid Diffusivity (Ft2/sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE LDIFENSI(LDSI,LDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDIFENSI
!MS$ ATTRIBUTES ALIAS:'_LDIFENSI'::LDIFENSI
!MS$ ATTRIBUTES REFERENCE::LDSI,LDENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION LDENG, LDSI
        LDSI = LDENG / (3.2808D0**2)                       
      END
 
!CC*******************************************************************


       
