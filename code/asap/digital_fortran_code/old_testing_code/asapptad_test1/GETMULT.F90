!C****************************************************************
!CC
!CC                        GETMULT.F
!CC
!CC Description:  This subroutine will calculate the multiple that
!CC               will result in VQ when it is multiplied by VQMIN.
!CC
!CC Output Variable:
!CC    XMPL =     Multiple of VQMIN to get VQ
!CC
!CC Input Variables:
!CC    VQ =       Air to water ratio (dimensionless)
!CC    VQMIN =    Minimum air to water ratio (dimensionless)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C****************************************************************

      SUBROUTINE GETMULT(XMPL,VQ,VQMIN)
!C  ATTRIBUTES DLLEXPORT, STDCALL::GETMULT
!C  ATTRIBUTES ALIAS:'_GETMULT@12':: GETMULT
!C  ATTRIBUTES REFERENCE::XMPL,VQ,VQMIN

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION XMPL,VQ,VQMIN

         XMPL = VQ/VQMIN

      END

!C****************************************************************

