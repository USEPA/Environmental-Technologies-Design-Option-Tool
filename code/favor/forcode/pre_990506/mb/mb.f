CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C  PROGRAM FOR CALCULATING SUBSTRATE AND BIOMASS CONCENTRATIONS     C
C  IN A SERIES OF CSTRS, WITH OR WITHOUT STEP FEED                  C
C                                                                   C
C  BASED ON EQUATIONS 5-14 AND 5-15, PAGE 352, DAVIS & CORNWELL,    C 
C  INTRODUCTION TO ENVIRONMENTAL ENGINEERING, MCGRAW-HILL, 1991.    C
C                                                                   C
C  UNITS ARE CONSISTENT WITHIN CODE                                 C
C                                                                   C
C  PARTIAL LIST OF VARIABLE DEFINITIONS                             C
C  --------------------------------------------------------------   C
C  ITERMX = MAXIMUM NUMBER OF ITERATIONS                            C
C  KD = BACTERIAL DECAY RATE (1/day)                                C
C  KS = HALF VELOCITY CONSTANT (mg BOD5/L)                          C
C  MUM = MAXIMUM GROWTH RATE CONSTANT (1/day)                       C
C  NN = NUMBER OF CSTRS                                             C
C  NSF = STEP FEED OPTION: NSF=0: NO STEP FEED; NSF=1: STEP FEED    C
C  Q0 = PLANT FLOWRATE (L/day)                                      C
C  QR = RECYCLE FLOWRATE (L/day)                                    C
C  QT = Q0+QR=TOTAL FLOW RATE (L/day)                               C
C  RESMX = ITERATION CONVERGENCE CRITERIA                           C
C  S0 = INFLUENT SUBSTRATE CONCENTRATION (mg BOD5/L)                C
C  S(N) = SUBSTRATE CONCN. IN CSTR N (mg BOD5/L)                    C
C  VT = TOTAL AERATION BASIN VOLUME (L)                             C
C  XR = BIOMASS CONCN IN RECYCLE STREAM (mg VSS/L)                  C
C  X(N) = BIOMASS CONCN IN CSTR N (mg VSS/L)                        C
C  Y = YIELD COEFFCIENT (mg VSS/mg BOD5)                            C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      IMPLICIT REAL*8 (A-H,O-Z)
      REAL*8 KX,KD,KS,MUM
      DIMENSION X(20),S(20),XOLD(20),SOLD(20),QN1(20),QFN(20),QN(20)

C.....INPUT PARAMETERS

C.....NUMBER OF CSTRS
      NN = 2
C.....STEP FEED OPTION: NSF=0: NO STEP FEED; NSF=1: STEP FEED
      NSF = 0
C.....PLANT FLOWRATE
      Q0 = 3.81D5
C.....RECYCLE FLOWRATE
      QR = 3.048D5
C.....WASTAGE FLOWRATE
      QW = 3.D3
C.....TOTAL AERATION BASIN VOLUME 
      VT = 6.D4
C.....MAXIMUM GROWTH RATE CONSTANT
      MUM = 3.D0
C.....HALF VELOCITY CONSTANT
      KS = 60.D0
C.....BACTERIAL DECAY RATE
      KD = 0.06D0
C.....YIELD COEFFCIENT
      Y = 0.6D0
C.....INFLUENT SUBSTRATE CONCENTRATION
      S0 = 84.D0
C.....MAXIMUM NUMBER OF ITERATIONS
      ITERMX = 1000
C.....ITERATION CONVERGENCE CRITERIA
      RESMX = 1.D-6

C.....INITIAL GUESSES
      XR = 2.D0*Y*S0
      S(1) = S0
      IF (NN.GT.2) THEN 
         DO 300 N=2,NN-1
            S(N) = 0.5D0*S0
  300    CONTINUE
      ENDIF
      IF (NN.GT.1) S(NN) = 0.2D0*S0

C.....CALCULATED PARAMETERS      
      QT = Q0 + QR
      VN = VT/DFLOAT(NN)
      IF (NN.EQ.1) NSF=0

C.....INITIALIZE
      ITER=0
      DO 100 N=1,NN
      X(N) = 0.D0
  100 S(N) = 0.D0

C.....BEGIN ITERATION LOOP

  400 CONTINUE

      ITER=ITER+1

      DO 450 N=1,NN
      XOLD(N)=X(N)
  450 SOLD(N)=S(N)

      ST = ((Q0*S0)+(QR*S(NN)))/QT
      IF(ITER.GT.1) XR = X(NN)*QT/(QW+QR)
      XT = QR*XR/QT

C.....N=1

      N = 1
      IF (NSF.EQ.1) QFN(1) = QT/DFLOAT(NN)
      IF (NSF.EQ.0) QFN(1) = QT 
      IF (NSF.EQ.1) QN(1) = DFLOAT(N)*QT/DFLOAT(NN)
      IF (NSF.EQ.0) QN(1) = QT
      XTX = XT
      STX = ST
      KX = (MUM*S(1)/(KS+S(1)))-KD
      X(1) = (QFN(1)*XTX)/(QN(1)-(VN*KX))
      KS = MUM*X(1)/(Y*(KS+S(1)))
      S(1) = (QFN(1)*STX)/(QN(1)+(VN*KS))
 
C.....BEGIN N LOOP
      IF (NN.GT.2) THEN

         DO 500 N=2,NN-1  

            QN1(N) = DFLOAT(N-1)*QT/DFLOAT(NN)
            IF (NSF.EQ.1) QFN(N) = QT/DFLOAT(NN)
            IF (NSF.EQ.0) QFN(N) = QT - QN1(N)
            IF (NSF.EQ.1) QN(N) = DFLOAT(N)*QT/DFLOAT(NN)
            IF (NSF.EQ.0) QN(N) = QT
            IF (NSF.EQ.1) XTX = XT
            IF (NSF.EQ.0) XTX = X(N-1)
            IF (NSF.EQ.1) STX = ST
            IF (NSF.EQ.0) STX = S(N-1)
            KX = (MUM*S(N)/(KS+S(N)))-KD
            X(N) = ((QN1(N)*X(N-1))+(QFN(N)*XTX))/(QN(N)-(VN*KX))
            KS = MUM*X(N)/(Y*(KS+S(N)))
            S(N) = ((QN1(N)*S(N-1))+(QFN(N)*STX))/(QN(N)+(VN*KS))
   
  500    CONTINUE

      ENDIF

C.....N=NN

      IF (NN.GT.1) THEN

         N = NN
         QN1(N) = DFLOAT(N-1)*QT/DFLOAT(NN)
         IF (NSF.EQ.1) QFN(N) = QT/DFLOAT(NN)
         IF (NSF.EQ.0) QFN(N) = QT - QN1(N)
         IF (NSF.EQ.1) QN(N) = DFLOAT(N)*QT/DFLOAT(NN)
         IF (NSF.EQ.0) QN(N) = QT
         IF (NSF.EQ.1) XTX = XT
         IF (NSF.EQ.0) XTX = X(N-1)
         IF (NSF.EQ.1) STX = ST
         IF (NSF.EQ.0) STX = S(N-1)
         KX = (MUM*S(N)/(KS+S(N)))-KD
         X(N) = ((QN1(N)*X(N-1))+(QFN(N)*XTX))/(QN(N)-(VN*KX))
         KS = MUM*X(N)/(Y*(KS+S(N)))
         S(N) = ((QN1(N)*S(N-1))+(QFN(N)*STX))/(QN(N)+(VN*KS))

      ENDIF
   
C.....CALCULATE RESIDUAL AND CHECK

      TRES = 0.D0
      DO 700 N=1,NN
      RES = DABS((SOLD(N)-S(N))/S(N)) 
     1    + DABS((XOLD(N)-X(N))/X(N))
  700 TRES = TRES + RES

C.....WRITE RESULTS

      IF ((TRES.GT.RESMX).AND.(ITER.LT.ITERMX)) GO TO 400
      IF ((TRES.GT.RESMX).AND.(ITER.EQ.ITERMX)) THEN
         WRITE (6,2000)
      ENDIF

c      WRITE (6,2010) ITER,TRES
c  750 WRITE (6,2030) XR
      DO 800 N=1,NN
c  800 WRITE (6,2020) N,X(N),S(N)
  800 WRITE (6,2030) X(N)

 2000 FORMAT (' EXCEEDED ITERMX WITHOUT CONVERGING')
 2010 FORMAT (I4,E12.5)
 2020 FORMAT (I4,2E12.5)
 2030 FORMAT (E12.5)

      STOP
      END
 
