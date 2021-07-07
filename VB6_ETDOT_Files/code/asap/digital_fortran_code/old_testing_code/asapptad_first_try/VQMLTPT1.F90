!C*************************************************************
!CC
!CC                     VQMLTPT1
!CC
!CC Description:  This subroutine will calculate air to water
!CC               ratio, given a value for minimum air to
!CC               water ratio, and a multiple to achieve
!CC               air to water ratio.
!CC
!CC Output Variable:
!CC    VQ =       Air to water ratio (dimensionless)
!CC
!CC Input Variables:
!CC    VQMIN =    Minimum air to water ratio (dimensionless)
!CC    XMPL =     Multiple of VQMIN to get VQ
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C*************************************************************

      SUBROUTINE VQMLTPT1(VQ,VQMIN,XMPL)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VQMLTPT1
!MS$ ATTRIBUTES ALIAS:'_VQMLTPT1@12':: VQMLTPT1
!MS$ ATTRIBUTES REFERENCE::VQ,VQMIN,XMPL

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VQ,VQMIN,XMPL

         VQ=XMPL*VQMIN           

      END

!C*************************************************************

