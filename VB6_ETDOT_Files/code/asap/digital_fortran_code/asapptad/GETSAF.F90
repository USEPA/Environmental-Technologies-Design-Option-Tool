!C***************************************************************
!CC
!CC                            GETSAF
!CC               FINDING THE KLa SAFETY FACTOR
!CC
!CC Description:  This subroutine will calculate the KLa safety
!CC               factor, given the Onda mass transfer coefficient
!CC               and the design mass transfer coefficient.
!CC
!CC Output Variable:
!CC    SAFFAC =   Safety factor on KLa (dimensionless)
!CC
!CC Input Variable:
!CC    KLAOND =   KLA calculated with the Onda correlation (1/sec)
!CC    KLASAF =   Design overall mass transfer coeffient (1/sec)
!CC
!CC
!CC History:  Subroutine written by David R. Hokanson (1/3/94)
!CC
!C***************************************************************

      SUBROUTINE GETSAF(SAFFAC,KLAOND,KLASAF)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETSAF
!MS$ ATTRIBUTES ALIAS:'_GETSAF':: GETSAF
!MS$ ATTRIBUTES REFERENCE::SAFFAC,KLAOND,KLASAF

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION KLASAF,KLAOND,SAFFAC 

         SAFFAC = KLASAF / KLAOND 

      END

!C***************************************************************

