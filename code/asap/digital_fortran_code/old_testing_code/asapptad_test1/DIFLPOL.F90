!C*************************************************************
!CC
!CC                    DIFLPOL
!CC
!CC Description:  This subroutine will calculate liquid diffusivity
!CC               for compounds with molecular weight > 1000.  It
!CC               uses the method of Polson, 1950.
!CC
!CC Output Variable:
!CC    DIFL =     Liquid diffusivity (m^2/sec)
!CC
!CC Input Variable:
!CC    MW =       Molecular weight of compound
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C*************************************************************

      SUBROUTINE DIFLPOL(DIFL,MW)
!C  ATTRIBUTES DLLEXPORT, STDCALL::DIFLPOL
!C  ATTRIBUTES ALIAS:'_DIFLPOL@8':: DIFLPOL
!C  ATTRIBUTES REFERENCE::DIFL,MW

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFL,MW

         DIFL = (2.74D-5)*(MW**(-1.0D0/3.0D0))            

      END

!C*************************************************************

