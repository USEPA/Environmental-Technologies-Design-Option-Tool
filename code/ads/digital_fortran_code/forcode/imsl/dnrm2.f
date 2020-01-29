C-----------------------------------------------------------------------
C  IMSL Name:  DNRM2 (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Compute the Euclidean length or L2 norm of a
C              double-precision vector.
C
C  Usage:      DNRM2(N, DX, INCX)
C
C  Arguments:
C     N      - Length of vector X.  (Input)
C     DX     - Double precision vector of length N*INCX.  (Input)
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be DX(1+(I-1)*INCX). INCX must be
C              greater than zero.
C     DNRM2  - Double precision square root of the sum from I=1 to N of
C              X(I)**2.  (Output)
C              X(I) refers to a specific element of DX. See INCX
C              argument description.
C
C  GAMS:       D1a3b
C
C  Chapters:   MATH/LIBRARY Basic Matrix/Vector Operations
C              STAT/LIBRARY Mathematical Support
C
C  Copyright:  1986 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      DOUBLE PRECISION FUNCTION DNRM2 (N, DX, INCX, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      DOUBLE PRECISION DX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, J, NEXT, NN
      DOUBLE PRECISION HITEST, SUM, XMAX
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      DOUBLE PRECISION CUTHI, CUTLO, ONE, ZERO
      SAVE       CUTHI, CUTLO, ONE, ZERO
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS,DSQRT
      INTRINSIC  DABS, DSQRT
      DOUBLE PRECISION DABS, DSQRT
C
      DATA ZERO/0.0D0/, ONE/1.0D0/
      DATA CUTLO/8.232D-11/, CUTHI/1.304D19/
C
      IF (N .GT. 0) GO TO 10
      DNRM2 = ZERO
      GO TO 140
C
   10 ASSIGN 30 TO NEXT
      SUM = ZERO
      NN = N*INCX
C                                  BEGIN MAIN LOOP
      I = 1
   20 GO TO NEXT, (30, 40, 70, 80)
   30 IF (DABS(DX(I)) .GT. CUTLO) GO TO 110
      ASSIGN 40 TO NEXT
      XMAX = ZERO
C                                  PHASE 1. SUM IS ZERO
   40 IF (DX(I) .EQ. ZERO) GO TO 130
      IF (DABS(DX(I)) .GT. CUTLO) GO TO 110
C                                  PREPARE FOR PHASE 2.
      ASSIGN 70 TO NEXT
      GO TO 60
C                                  PREPARE FOR PHASE 4.
   50 I = J
      ASSIGN 80 TO NEXT
      SUM = (SUM/DX(I))/DX(I)
   60 XMAX = DABS(DX(I))
      GO TO 90
C                                  PHASE 2. SUM IS SMALL. SCALE TO
C                                  AVOID DESTRUCTIVE UNDERFLOW.
   70 IF (DABS(DX(I)) .GT. CUTLO) GO TO 100
C                                  COMMON CODE FOR PHASES 2 AND 4. IN
C                                  PHASE 4 SUM IS LARGE. SCALE TO
C                                  AVOID OVERFLOW.
   80 IF (DABS(DX(I)) .LE. XMAX) GO TO 90
      SUM = ONE + SUM*(XMAX/DX(I))**2
      XMAX = DABS(DX(I))
      GO TO 130
C
   90 SUM = SUM + (DX(I)/XMAX)**2
      GO TO 130
C                                  PREPARE FOR PHASE 3.
  100 SUM = (SUM*XMAX)*XMAX
C                                  FOR REAL OR D.P. SET HITEST =
C                                  CUTHI/N FOR COMPLEX SET HITEST =
C                                  CUTHI/(2*N)
  110 HITEST = CUTHI/N
C                                  PHASE 3. SUM IS MID-RANGE. NO
C                                  SCALING.
      DO 120  J=I, NN, INCX
         IF (DABS(DX(J)) .GE. HITEST) GO TO 50
  120 SUM = SUM + DX(J)**2
      DNRM2 = DSQRT(SUM)
      GO TO 140
C
  130 CONTINUE
      I = I + INCX
      IF (I .LE. NN) GO TO 20
C                                  END OF MAIN LOOP. COMPUTE SQUARE
C                                  ROOT AND ADJUST FOR SCALING.
      DNRM2 = XMAX*DSQRT(SUM)
  140 CONTINUE
      RETURN
      END
