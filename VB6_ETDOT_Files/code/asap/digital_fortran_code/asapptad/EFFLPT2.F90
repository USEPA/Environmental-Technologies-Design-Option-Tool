!C***************************************************************
!CC
!CC                       EFFLPT2
!CC
!CC Description:  This subroutine will calculate the effluent
!CC               concentration for a given compound specified
!CC               earlier.
!CC
!CC Output Variables:
!CC    CE =       Effluent concentration (ug/L)
!CC
!CC Input Variables:
!CC    VQ =       Air to water ratio (dimensionless)
!CC    HC =       Henry's constant (dimensionless)
!CC    QW =       Water flow rate (m^3/sec)
!CC    AREA =     Tower area (m^2)
!CC    HLL =      Tower length (m)
!CC    KLA =      Overall mass transfer coefficient (1/sec)
!CC    CI =       Influent concentration (ug/L)
!CC
!CC Variables Internal to Subroutine EFFLPT2
!CC    RR =       Stripping Factor
!CC    QWA =      Volumetric water loading rate (m^3/m^2/sec)
!CC    BB =       Variable used to simplify calculation
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE EFFLPT2(CE,VQ,HC,QW,AREA,HLL,KLA,CI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::EFFLPT2
!MS$ ATTRIBUTES ALIAS:'_EFFLPT2@32':: EFFLPT2
!MS$ ATTRIBUTES REFERENCE::CE,VQ,HC,QW,AREA,HLL,KLA,CI
      
         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION CE,VQ,HC,QW,AREA,HLL,KLA,CI,RR,QWA,BB

         RR = VQ*HC
         QWA = QW/AREA
         BB = (HLL*KLA*(RR-1))/(QWA*RR)   
         CE = (CI*(RR-1))/(RR*(EXP(BB))-1)

      END

!C***************************************************************

