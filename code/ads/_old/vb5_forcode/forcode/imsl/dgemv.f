C-----------------------------------------------------------------------
C  IMSL Name:  DGEMV  (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    July 16, 1986
C
C  Purpose:    Perform the matrix-vector multiplication
C              y = alpha*A*x + beta*y or y = alpha*A'*x + beta*y,
C              all double precision.
C
C  Usage:      CALL DGEMV (TRANS, M, N, ALPHA, A, LDA, X, INCX, BETA, Y,
C                          INCY)
C
C  Arguments:
C     TRANS  - Character specifing the operation to be performed.
C              (Input)
C                 TRANS               Operation
C              'N' or 'n'      y = alpha*A*x + beta*y
C              'T' or 't'      y = alpha*A'*x + beta*y
C              'C' or 'c'      y = alpha*A'*x + beta*y
C     M      - Number of rows in A.  (Input)
C     N      - Number of columns in A.  (Input)
C     ALPHA  - Double precision scalar.  (Input)
C     A      - Array of size M by N.  (Input)
C     LDA    - Leading dimension of A exactly as specified in the
C              calling routine.  (Input)
C     X      - Double precision vector of length (N-1)*IABS(INCX)+1 when
C              TRANS is 'N' or 'n' and of length (M-1)*IABS(INCX)+1
C              otherwise.  (Input)
C     INCX   - Displacement between elements of X.  (Input)
C     BETA   - Double precision scalar.  (Input)
C              When BETA is zero, Y is not referenced.
C     Y      - Double precision vector of length (N-1)*IABS(INCY)+1 when
C              TRANS is 'M' or 'm' and of length (M-1)*IABS(INCY)+1
C              otherwise.  (Input/Output)
C     INCY   - Displacement between elements of Y.  (Input)
C
C  GAMS:       D1b
C
C  Chapter:    MATH/LIBRARY Basic Matrix/Vector Operations
C
C  Copyright:  1986 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DGEMV (TRANS, M, N, ALPHA, A, LDA, X, INCX, BETA, Y,
     &                  INCY)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    M, N, LDA, INCX, INCY
      DOUBLE PRECISION ALPHA, BETA, X(*), Y(*)
      CHARACTER  TRANS*1
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IX, IY, J, KY, LENX, LENY
C                                  SPECIFICATIONS FOR SPECIAL CASES
      DOUBLE PRECISION A(*)
      EXTERNAL   DAXPY, DDOT
      INTEGER    IA, KX
      DOUBLE PRECISION DDOT
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  IABS
      INTRINSIC  IABS
      INTEGER    IABS
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   DSCAL, DSET
C                                  Quick return if possible
      IF (M.EQ.0 .OR. N.EQ.0 .OR. (ALPHA.EQ.0.0) .AND. (BETA.EQ.1.0))
     &    GO TO 9000
C
      IF (TRANS.EQ.'N' .OR. TRANS.EQ.'n') THEN
         LENX = N
         LENY = M
      ELSE
         LENX = M
         LENY = N
      END IF
C
      IX = 1
      IY = 1
      IF (INCX .LT. 0) IX = (-LENX+1)*INCX + 1
      IF (INCY .LT. 0) IY = (-LENY+1)*INCY + 1
C
      IF (BETA .EQ. 1) THEN
      ELSE IF (INCY .EQ. 0) THEN
         IF (BETA .EQ. 0.0) THEN
            Y(1) = 0.0
         ELSE
            Y(1) = BETA**LENY*Y(1)
         END IF
      ELSE IF (BETA .EQ. 0.0) THEN
         CALL DSET (LENY, 0.0D0, Y, IABS(INCY))
      ELSE
         CALL DSCAL (LENY, BETA, Y, IABS(INCY))
      END IF
C
      IF (ALPHA .EQ. 0.0) GO TO 9000
C                                  Not transpose
      IF (TRANS.EQ.'N' .OR. TRANS.EQ.'n') THEN
         KX = IX
         DO 10  I=1, N
            CALL DAXPY (M, ALPHA*X(KX), A(LDA*(I-1)+1), 1, Y, INCY)
            KX = KX + INCX
   10    CONTINUE
      ELSE
C                                  Transpose
         KY = IY
         DO 20  I=1, N
            Y(KY) = Y(KY) + ALPHA*DDOT(M,A(LDA*(I-1)+1),1,X,INCX)
            KY = KY + INCY
   20    CONTINUE
      END IF
C
 9000 RETURN
      END
