!CC*******************************************************************
!CC
!CC                                GDIFENSI
!CC          CONVERT GAS DIFFUSIVITY FROM Ft2/sec to m2/sec
!CC
!CC
!CC Description:  This SUBROUTINE will convert gas diffusivity
!CC               from units of Ft2/sec to m2/sec
!CC
!CC Output Variables:
!CC    GDSI =     Gas Diffusivity (m2/sec)
!CC
!CC Input Variables:
!CC    GDENG =    Gas Diffusivity (Ft2/sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE GDIFENSI(GDSI,GDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GDIFENSI
!MS$ ATTRIBUTES ALIAS:'_GDIFENSI'::GDIFENSI
!MS$ ATTRIBUTES REFERENCE::GDSI,GDENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION GDENG, GDSI
        GDSI = GDENG / (3.2808D0**2)                       
      END
 
!CC*******************************************************************


       
