!CC****************************************************************
!CC
!CC                             GETGAM
!CC         CALCULATE INFINITE DILUTION ACTIVITY COEFFICIENT
!CC
!CC Description:  This subroutine will calculate the infinite dilution
!CC               activity coefficient at a given temperature from
!CC               UNIFAC.
!CC
!CC Output Variable:
!CC    GAMMA =    Infinite dilution activity coefficent (-)
!CC
!CC Input Variables:
!CC    NC =
!CC    NG =
!CC    TT =       Operating temperature (K)
!CC    NDIF =
!CC    XX =
!CC    ACT =
!CC    DACT =
!CC    TACT =
!CC
!CC Variables Internal to Subroutine GETGAM:
!CC    NACT =
!CC
!CC****************************************************************

      SUBROUTINE GETGAM(GAMMA,NC,NG,TT,NDIF,XX,ACT,DACT,TACT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETGAM
!MS$ ATTRIBUTES ALIAS:'_GETGAM@36':: GETGAM
!MS$ ATTRIBUTES REFERENCE::GAMMA,NC,NG,TT,NDIF,XX,ACT,DACT,TACT

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         PARAMETER (ND=10)
         DIMENSION XX(10),ACT(ND),DACT(ND),TACT(ND)       
         DOUBLE PRECISION GAMMA
        
         NACT = 0
         CALL PARMS(NC,NG,TT)
         CALL UNIMOD(NDIF,NACT,NC,NG,TT,XX,ACT,DACT,TACT)  
         GAMMA = ACT(2)

         RETURN

      END
 
!CC****************************************************************

