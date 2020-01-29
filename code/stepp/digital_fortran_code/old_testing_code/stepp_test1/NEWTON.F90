!CC*********************************************************************
!CC
!CC     NEWTON-RAPHSON ALGORITHM (GAUSS-JORDAN MAXIMUM PIVOT STRATEGY)
!CC   ------------------------------------------------------------------
!CC   THIS IS A SUBROUTINE TO IMPLEMENT THE NEWTON-RAPHSON ALGORITHM FOR
!CC   SOLVING SYSTEMS OF NONLINEAR ALGEBRAIC EQUATIONS.  A VARIATION OF
!CC   THE GAUSS-JORDAN MAXIMUM PIVOT STRATEGY IS EMPLOYED TO DETERMINE
!CC   THE INVERSE OF THE JACOBIAN MATRIX.  THE CORRECTION FACTORS ARE
!CC   CALCULATED IN AN ITERATIVE MANNER TO BRING THE ADJUSTABLE
!CC   VARIABLES WITHIN A SPECIFIED TOLERANCE.
!CC
!CC                   MM = NUMBER OF COLUMNS IN MATRIX C
!CC                   NN = NUMBER OF ROWS IN MATRIX C
!CC
!CC Output Variables:
!CC    XX =
!CC    FF =
!CC    IERR =      Error flag from this routine
!CC
!CC Input Variables:
!CC    NN =
!CC    TT =        Temperature of the calculation (K)
!CC    NG =
!CC    MAXIT =     Maximum number of iterations
!CC    TOL =       Tolerance
!CC    XGUESS =    Initial guesses array
!CC
!CC Authors:  M. Miller and T. Rogers (4/5/94)
!CC
!CC*********************************************************************

      SUBROUTINE NEWTON (NN,TT,NG,MAXIT,TOL,XGUESS,XX,FF,IERR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::NEWTON
!MS$ ATTRIBUTES ALIAS:'_NEWTON@36':: NEWTON
!MS$ ATTRIBUTES REFERENCE::NN,TT,NG,MAXIT,TOL,XGUESS,XX,FF,IERR

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      DIMENSION CC(2,3),IR(2),DX(2),XOLD(2),XX(2)
      DIMENSION FOLD(2),FF(2),XGUESS(2),X1(10),X2(10)
      DIMENSION ACT1(10),DACT1(10,10),TACT1(10)
      DIMENSION ACT2(10),DACT2(10,10),TACT2(10)

      COMMON /ERR/ ERRMAT(30),ERRNUM


!CC    -- SET WIDTH OF JACOBIAN MATRIX

      MM = NN + 1

!CC    -- INITIALIZATION (CC, DX, XX, XOLD, FOLD)

      DO 30 L=1,NN

           XX(L) = XGUESS(L)
           XOLD(L) = XGUESS(L)
           FF(L) = 0.0D0
           FOLD(L) = 0.0D0
           DX(L) = 0.0D0

           DO 20 LL=1,MM

                CC(L,LL) = 0.0D0

  20       CONTINUE

  30  CONTINUE

      NDIF = 1
      NACT = 0

      CALL PARMS (NN,NG,TT)

!CC    -- START OF NEWTON-RAPHSON ITERATION LOOP

      IMAX = MAXIT + 1

      DO 180 ITER=1,IMAX

           INDEX = ITER - 1

!CC    -- EVALUATE FUNCTION VECTOR (FF)

           X1(1) = XX(1)
           X1(2) = 1.0D0 - X1(1)
           X2(1) = XX(2)
           X2(2) = 1.0D0 - X2(1)

           CALL UNIMOD (NDIF,NACT,NN,NG,TT,X1,ACT1,DACT1,TACT1)
           CALL UNIMOD (NDIF,NACT,NN,NG,TT,X2,ACT2,DACT2,TACT2)

!CC    -- CALCULATE ROOT-MEAN-SQUARE ERROR (RMSE)

           LOGIC = 0
           RMSE1 = 0.0D0
           RMSE2 = 0.0D0
           RMSE3 = 0.0D0

           DO 40 J=1,NN
 
                FF(J) = X1(J)*ACT1(J) - X2(J)*ACT2(J)

                IF (DABS(FF(J)).GE.TOL) LOGIC = -1

                DIF1 = FF(J) - FOLD(J)
                DIF2 = XX(J) - XOLD(J)
                RMSE1 = RMSE1 + FF(J)**2
                RMSE2 = RMSE2 + (DIF1)**2
                RMSE3 = RMSE3 + (DIF2)**2

  40       CONTINUE

           IF (LOGIC.EQ.INDEX) GOTO 200

           RMSE1 = DSQRT(RMSE1/DBLE(NN))
           RMSE2 = DSQRT(RMSE2/DBLE(NN))
           RMSE3 = DSQRT(RMSE3/DBLE(NN))

!CC    -- TEST FOR CONVERGENCE OF SOLUTION

           IF (RMSE1.GE.TOL) GOTO 50

           IF (RMSE2.GE.TOL) GOTO 50

           IF (RMSE3.GE.TOL) GOTO 50

           IF (LOGIC.EQ.-1)  GOTO 50

           GOTO 200

!CC    -- SAVE PREVIOUS ITERATION

  50       DO 60 I=1,NN

                FOLD(I) = FF(I)
                XOLD(I) = XX(I)

  60       CONTINUE

!CC    -- LOAD PARTIAL DERIVATIVES IN JACOBIAN

           CC(1,1) =  ACT1(1) + XX(1)*DACT1(1,1)
           CC(1,2) = -ACT2(1) - XX(2)*DACT2(1,1)
           CC(2,1) = -ACT1(2) + (1.0D0 - XX(1))*DACT1(2,1)
           CC(2,2) =  ACT2(2) - (1.0D0 - XX(2))*DACT2(2,1)

!CC    -- FINISH LOADING "CC" MATRIX WITH "FF" VECTOR

           DO 70 I=1,NN

                CC(I,MM) = -FF(I)

  70       CONTINUE

!CC    -- GAUSS-JORDAN ALGORITHM
!CC    -- INITIALIZE ALL VECTORS AND MATRICES

           DO 80 I=1,NN

                DX(I) = 0.0D0
                IR(I) = 0
                JJ = 0
                JM = 0

  80       CONTINUE

           DO 140 K=1,NN

                PK = 0.0D0

!CC    -- LOCATE PIVOT ELEMENT

                DO 100 I=1,NN
              
                     IF (I.EQ.IR(I)) GOTO 100

                     DO 90 IK=1,NN

                          PP = DABS(CC(I,IK))

                          IF (PP.LT.PK) GOTO 90

                          PK = PP
                          JJ = I
                          JM = IK

  90                 CONTINUE

 100            CONTINUE
        
                IR(JJ) = JJ

!CC    -- NORMALIZATION STEP

                DO 110 JR=1,MM

                     IF (JM.EQ.JR) GOTO 110

                     IF (DABS(CC(JJ,JM)).LE.1.0D-25) THEN

                          IERR = -2
                          CALL ERROR (ERRMAT,ERRNUM,23)
                          RETURN
      
                     END IF 

                     CC(JJ,JR) = CC(JJ,JR)/CC(JJ,JM)

 110            CONTINUE

                CC(JJ,JM) = 1.0D0

!CC    -- REDUCTION STEP

                DO 130 I=1,NN

                     IF (I.EQ.JJ) GOTO 130

                     DO 120 JR=1,MM

                          IF (JR.EQ.JM) GOTO 120

                          CC(I,JR) = CC(I,JR) - CC(I,JM) * CC(JJ,JR)

 120                 CONTINUE

                     CC(I,JM) = 0.0D0

 130            CONTINUE

 140       CONTINUE

!CC    -- END OF GAUSS-JORDAN MAXIMUM PIVOT ROUTINE
!CC    -- RECOVER THE SOLUTION VECTOR

           DO 160 I=1,NN

                DO 150 J=1,NN

                     IF((CC(I,J).LT.1).OR.(CC(I,J).GT.1)) GOTO 150

                     DX(J) = CC(I,MM)

 150            CONTINUE

 160       CONTINUE

!CC    -- CORRECT ELEMENTS OF THE "XX" VECTOR

           DO 170 I=1,NN

                 XX(I) = XX(I) + DX(I)

                 IF(XX(I).LT.0.D0) XX(I) = 0.0D0

                 IF(XX(I).GT.1.D0) XX(I) = 1.0D0

 170       CONTINUE

 180  CONTINUE

!CC    -- END OF ITERATION LOOP

!CC    -- CONVERGENCE FAILURE MESSAGE (MAXIT REACHED)

 190  IERR = -1
      CALL ERROR (ERRMAT,ERRNUM,22)

 200  END


