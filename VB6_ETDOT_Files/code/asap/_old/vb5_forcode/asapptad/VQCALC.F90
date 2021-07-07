!C***************************************************************
!CC
!CC                            VQCALC
!CC         CALCULATION OF THE AIR TO WATER RATIO
!CC
!CC Description:  This subroutine will calculate the air to water
!CC               ratio, given air flow rate and water flow rate
!CC
!CC Output Variable:
!CC    VQ =       Air to water ratio (dimensionless)
!CC
!CC Input Variable:
!CC    QA =       Air flow rate (m^3/sec)
!CC    QW =       Water flow rate (m^3/sec)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE VQCALC(VQ,QA,QW)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VQCALC
!MS$ ATTRIBUTES ALIAS:'_VQCALC':: VQCALC
!MS$ ATTRIBUTES REFERENCE::VQ,QA,QW

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION QA,VQ,QW

         VQ = QA/QW

      END

!C***************************************************************

