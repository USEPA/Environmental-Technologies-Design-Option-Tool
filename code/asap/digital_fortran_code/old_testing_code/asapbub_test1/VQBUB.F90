!C***************************************************************
!CC
!CC                            VQBUB
!CC         CALCULATION OF THE AIR TO WATER RATIO
!CC
!CC Description:  This subroutine will calculate the air to water
!CC               ratio, given air flow rate to each tank and
!CC               water flow rate.
!CC
!CC Output Variable:
!CC    VQ =       Air to water ratio (dimensionless)
!CC
!CC Input Variable:
!CC    QA =       Air flow rate to each tank (m^3/sec)
!CC    QW =       Water flow rate (m^3/sec)
!CC
!C***************************************************************

      SUBROUTINE VQBUB(VQ,QA,QW)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VQBUB
!MS$ ATTRIBUTES ALIAS:'_VQBUB':: VQBUB
!MS$ ATTRIBUTES REFERENCE::VQ,QA,QW

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION QA,VQ,QW

         VQ = QA/QW

      END

!C***************************************************************

