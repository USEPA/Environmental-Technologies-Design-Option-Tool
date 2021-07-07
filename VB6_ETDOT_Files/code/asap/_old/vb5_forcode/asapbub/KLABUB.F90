!C***************************************************************
!CC
!CC                         KLABUB
!CC               FIND KLA FOR BUBBLE AERATION
!CC
!CC Description:  This subroutine finds KLa for a compound for
!CC               bubble aeration using two film theory and
!CC               mass transfer correlations.
!CC
!CC Output Variables:
!CC    KLA =      Compound mass transfer coefficient (1/sec)
!CC    N =        Exponent used in correlation
!CC    KGKL =     Ratio of gas-phase to liquid-phase mass transfer
!CC               coefficent (assumed constant and equal to 100
!CC               for bubble aeration) ** Find Source **
!CC
!CC Input Variables:
!CC    KLAO2 =    Oxygen mass transfer coeff. (1/sec)
!CC    DIFL =     Diffusivity of liquid water (m^2/sec)
!CC    DIFLO2 =   Diffusivity of oxygen (m^2/sec)
!CC    HC =       Henry's constant (dimensionless)
!CC
!C***************************************************************

      SUBROUTINE KLABUB(KLA,KLAO2,DIFL,DIFLO2,N,KGKL,HC) 
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KLABUB
!MS$ ATTRIBUTES ALIAS:'_KLABUB':: KLABUB
!MS$ ATTRIBUTES REFERENCE::KLA,KLAO2,DIFL,DIFLO2,N,KGKL,HC

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION KLA,KLAO2,DIFL,DIFLO2,N,KGKL,HC

         N = 0.6D0
         KGKL = 100.0D0
         KLA = KLAO2 * ((DIFL/DIFLO2)**N) * (1.0D0/(1.0D0+(1.0D0/KGKL/HC)))

      END

!C***************************************************************

