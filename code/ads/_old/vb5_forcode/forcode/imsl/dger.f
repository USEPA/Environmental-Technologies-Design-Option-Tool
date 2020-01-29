C-----------------------------------------------------------------------
C  IMSL Name:  DGER  (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    July 17, 1986
C
C  Purpose:    Perform the rank-one matrix update A = alpha*x*y' + A,
C              all double precision.
C
C  Usage:      CALL DGER (M, N, ALPHA, X, INCX, Y, INCY, A, LDA)
C
C  Arguments:
C     M      - Number of rows in A.  (Input)
C     N      - Number of columns in A.  (Input)
C     ALPHA  - Double precision scalar.  (Input)
C     X      - Double precision vector of length (M-1)*IABS(INCX)+1.
C              (Input)
C     INCX   - Displacement between elements of X.  (Input)
C     Y      - Double precision vector of length (N-1)*IABS(INCY)+1.
C              (Input)
C     INCY   - Displacement between elements of Y.  (Input)
C     A      - Double precision array of size M by N.  (Input/Output)
C              On input, A contains the matrix to be updated.
C              On output, A contains the updated matrix.
C     LDA    - Leading dimension of A exactly as specified in the
C              calling routine.  (Input)
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
      SUBROUTINE DGER (M, N, ALPHA, X, INCX, Y, INCY, A, LDA)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    M, N, INCX, INCY, LDA
      DOUBLE PRECISION ALPHA, X(*), Y(*)
      DOUBLE PRECISION A(*)
      INTEGER    I1X
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IY, J
C                                  SPECIFICATIONS FOR SPECIAL CASES
      EXTERNAL   DAXPY
C                                  Quick return if possible
      IF (M.EQ.0 .OR. N.EQ.0 .OR. ALPHA.EQ.0.0D0) GO TO 9000
C
      IY = 1
      IF (INCY .LT. 0) IY = (-N+1)*INCY + 1
C
      I1X = 1
      DO 10  J=1, N
         CALL DAXPY (M, ALPHA*Y(IY), X, INCX, A(I1X), 1)
         IY = IY + INCY
         I1X = I1X + LDA
   10 CONTINUE
C
 9000 RETURN
      END
