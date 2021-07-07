!C***************************************************************
!CC
!CC                         EFFLUENT
!CC
!CC Description:  This subroutine will calculate the effluent
!CC               concentrations out of each tank for both the
!CC               liquid phase and the gas phase.
!CC
!CC Output Variables:
!CC    CEXIT =    Array of liquid phase effluent concs.
!CC    YEXIT =    Array of gas phase effluent concs.
!CC
!CC Input Variables:
!CC    HC =       Henry's constant of compound (dimensionless)
!CC    CI =       Liquid phase influent conc. (ug/L)
!CC    VQ =       Air to water ratio (dimensionless)
!CC    NTANK =    No. of tanks
!CC    PHI =      Stanton No. (dimensionless)
!CC
!C***************************************************************

      SUBROUTINE EFFLBUB(CEXIT,YEXIT,HC,CI,VQ,NTANK,PHI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::EFFLBUB
!MS$ ATTRIBUTES ALIAS:'_EFFLBUB':: EFFLBUB
!MS$ ATTRIBUTES REFERENCE::CEXIT,YEXIT,HC,CI,VQ,NTANK,PHI
         
         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DIMENSION CEXIT(0:20),YEXIT(20)
         INTEGER NTANK
         DOUBLE PRECISION CEXIT,YEXIT,HC,CI,VQ

         CEXIT(0) = CI
         DO 1400, I=1,NTANK
            CEXIT(I) = CI/((1.0D0+VQ*HC*(1.0D0-EXP(-PHI)))**I)
            YEXIT(I) = (1.0D0/VQ)*(CEXIT(I-1)-CEXIT(I))
1400     CONTINUE

      END

!C***************************************************************

