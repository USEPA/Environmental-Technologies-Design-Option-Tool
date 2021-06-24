!C***************************************************************
!CC
!CC                            KLACOR
!CC     FINDING THE DESIGN OVERALL MASS TRANSFER COEFFICIENT
!CC
!CC Description:  This subroutine will calculate the design overall
!CC               mass transfer coefficient by adjusting the KLa
!CC               calculated by the Onda correlation with a safety
!CC               factor.
!CC
!CC Output Variable:
!CC    KLASAF =   Design overall mass transfer coeffient (1/sec)
!CC
!CC Input Variable:
!CC    KLAOND =   KLA calculated with the Onda correlation (1/sec)
!CC    SAFFAC =   Safety factor on KLa (dimensionless)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE KLACOR(KLASAF,KLAOND,SAFFAC)
!C  ATTRIBUTES DLLEXPORT, STDCALL::KLACOR
!C  ATTRIBUTES ALIAS:'_KLACOR@12':: KLACOR
!C  ATTRIBUTES REFERENCE::KLASAF,KLAOND,SAFFAC

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION KLASAF,KLAOND,SAFFAC 

         KLASAF = KLAOND * SAFFAC 

      END

!C***************************************************************

