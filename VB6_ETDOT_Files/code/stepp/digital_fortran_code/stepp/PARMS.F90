!CC***************************************************************************
!CC
!CC                                  PARMS
!CC                      CALCULATE PARAMETERS FOR UNIMOD
!CC
!CC Input Variables:
!CC    NC =
!CC    NG =
!CC    T =        Operating temperature (K)
!CC
!CC Authors:  M. Miller, T. Rogers, D. Hokanson (4/4/94)
!CC
!CC***************************************************************************

      SUBROUTINE PARMS (NC,NG,T)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PARMS
!MS$ ATTRIBUTES ALIAS:'_PARMS@12':: PARMS
!MS$ ATTRIBUTES REFERENCE::NC,NG,T

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      COMMON /UNI/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),R(10),P(10,10)

      DO 10 I=1,NG

           DO 10 J=1,NG
           
                TAU(I,J) = DEXP(-P(I,J)/T)

  10  CONTINUE

      DO 20 I=1,NC

           DO 20 K=1,NG
    
                S(K,I) = 0

                     DO 20 M=1,NG

                          S(K,I) = S(K,I)+QT(M,I)*TAU(M,K)
  20  CONTINUE

      DO 30 I=1,NC

           F(I) = 1

           DO 30 J=1,NG
       
                F(I) = F(I)+QT(J,I)*DLOG(S(J,I))

  30  CONTINUE

      END


