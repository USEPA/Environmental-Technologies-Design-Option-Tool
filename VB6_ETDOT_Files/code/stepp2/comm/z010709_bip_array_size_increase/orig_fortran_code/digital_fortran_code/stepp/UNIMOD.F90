!CC***************************************************************************
!CC
!CC                               UNIMOD
!CC                    CALCULATE ACTIVITY COEFFICIENTS
!CC
!CC Output Variables:
!CC    ACT =
!CC    DACT = ]
!CC    TACT =
!CC
!CC Input Variables;
!CC    NC =
!CC    NG =
!CC    T =        Operating temperature (K)
!CC    X =
!CC
!CC Authors:  M. Miller, T. Rogers, D. Hokanson
!CC
!CC***************************************************************************

      SUBROUTINE UNIMOD (NDIF,NACT,NC,NG,T,X,ACT,DACT,TACT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::UNIMOD
!MS$ ATTRIBUTES ALIAS:'_UNIMOD@36':: UNIMOD
!MS$ ATTRIBUTES REFERENCE::NDIF,NACT,NC,NG,T,X,ACT,DACT,TACT

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      COMMON /UNI/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),R(10),P(10,10)

      DIMENSION X(10),GAM(10),ACT(10),DACT(10,10),THETA(10)
      DIMENSION PHI(10),RI(10),QI(10),QIL(10),RIL(10)
      DIMENSION QID(10),ETAL(10),TACT(10),U(10,10),V(10,10)
      DIMENSION DETA(10),DS(10,10),ETA(10),TETAR(10),H3(10,10)

      ZCOORD = 10.0D0
      THETS = 0.0D0
      PHS = 0.0D0

      DO 10 I=1,NC

           THETA(I) = X(I)*Q(I)
           PHI(I) = R(I)*X(I)
           THETS = THETS+THETA(I)
           PHS = PHS+PHI(I)

  10  CONTINUE

      DO 20 I=1,NC

         THETA(I) = THETA(I)/THETS
         PHI(I) = PHI(I)/PHS
         RI(I) = R(I)/PHS
         RIL(I) = DLOG(RI(I))
         QI(I) = Q(I)/THETS
         QID(I) = 1.0D0-RI(I)/QI(I)
         QIL(I) = DLOG(QI(I))

  20  CONTINUE

      DO 30 I=1,NC

         XX = F(I)+Q(I)*(1.0D0-QIL(I))-RI(I)+RIL(I)
         XX = XX-(ZCOORD/2.0D0)*Q(I)*(QID(I)+RIL(I)-QIL(I))
         GAM(I) = XX

  30  CONTINUE

      DO 50 I=1,NG

         TETAR(I) = 0.0D0
         ETA(I) = 0.0D0

         DO 40 J=1,NC

            ETA(I) = ETA(I)+S(I,J)*X(J)
            TETAR(I) = TETAR(I)+QT(I,J)*X(J)

  40     CONTINUE

         ETAL(I) = DLOG(ETA(I))

  50  CONTINUE

      DO 70 I=1,NC

         DO 60 J=1,NG

            U(J,I) = S(J,I)/ETA(J)
            V(J,I) = U(J,I)*TETAR(J)
            GAM(I) = GAM(I)-V(J,I)-QT(J,I)*ETAL(J)

  60     CONTINUE

         ACT(I) = DEXP(GAM(I))
         IF(NACT.EQ.1) ACT(I) = ACT(I)*X(I)

  70  CONTINUE

      IF (NDIF.EQ.0) GOTO 160

      IF (NDIF.EQ.2) GOTO 110

      DO 90 I=1,NC

           DO 90 J=1,NC

                XX = Q(I)*QI(J)*(1.0D0-(ZCOORD/2.0D0)*QID(I)*QID(J))+(1.0D0-RI(I))*(1.0D0-RI(J))

                DO 80 K=1,NG
 
                     XX = XX+U(K,I)*(V(K,J)-QT(K,J))-U(K,J)*QT(K,I)

  80            CONTINUE

           DACT(I,J) = XX
           DACT(J,I) = XX

           IF (NACT.EQ.1) GOTO 90

           DACT(I,J) = DACT(I,J)*ACT(I)

           IF (J.EQ.I) GOTO 90

           DACT(J,I) = DACT(J,I)*ACT(J)

  90  CONTINUE

      IF (NACT.EQ.0) GOTO 110

      DO 100 I=1,NC
     
           DO 100 J=1,NC

                DACT(I,J) = ACT(I)*(DACT(I,J)-1)
      
                IF (J.EQ.I) DACT(I,J) = DACT(I,J)+DEXP(GAM(I))

 100  CONTINUE

 110  IF (NDIF.EQ.1) GOTO 160

      DO 130 K=1,NG

           DETA(K) = 0

      DO 130 I=1,NC

           DS(K,I) = 0

           DO 120 M=1,NG

                 IF (QT(M,I).EQ.0) GOTO 120

                 DS(K,I) = DS(K,I)-QT(M,I)*DLOG(TAU(M,K))*TAU(M,K)/T

 120       CONTINUE

           DETA(K) = DETA(K)+DS(K,I)*X(I)

 130  CONTINUE

      DO 150 I=1,NC

           TACT(I) = 0

           DO 140 K=1,NG

                H3(K,I) = (-S(K,I)*DETA(K)/ETA(K)+DS(K,I))/ETA(K)
                HH = H3(K,I)*(TETAR(K)-QT(K,I)*ETA(K)/S(K,I))
                TACT(I) = TACT(I)-HH

 140       CONTINUE

           TACT(I) = TACT(I)*ACT(I)

 150  CONTINUE

 160  END


