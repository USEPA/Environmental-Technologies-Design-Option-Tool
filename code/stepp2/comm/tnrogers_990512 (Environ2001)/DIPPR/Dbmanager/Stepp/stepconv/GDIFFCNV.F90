!CC********************************************************************
!CC
!CC                             GDIFFCNV
!CC          CONVERT GAS DIFFUSIVITY UNITS FROM m2/sec TO ft2/sec
!CC
!CC Description:  This SUBROUTINE will convert gas diffusivity from
!CC               units of Pa to units of psi.
!CC
!CC Output Variables:
!CC    GDIFENG =  Gas Diffusivity (ft2/sec)
!CC
!CC Input Variables:
!CC    GDIFSI =   Gas Diffusivity (m2/sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/16/94)
!CC
!CC********************************************************************

      SUBROUTINE GDIFFCNV(GDIFENG,GDIFSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GDIFFCNV
!MS$ ATTRIBUTES ALIAS:'_GDIFFCNV'::GDIFFCNV
!MS$ ATTRIBUTES REFERENCE::GDIFENG,GDIFSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION GDIFENG,GDIFSI

         GDIFENG = GDIFSI * (3.2808D0**2)

      END

!CC********************************************************************


