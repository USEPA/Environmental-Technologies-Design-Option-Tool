C-----------------------------------------------------------------------
C  IMSL Name:  DSCAL (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Multiply a vector by a scalar, y = ay, both double
C              precision.
C
C  Usage:      CALL DSCAL (N, DA, DX, INCX)
C
C  Arguments:
C     N      - Length of vector X.  (Input)
C     DA     - Double precision scalar.  (Input)
C     DX     - Double precision vector of length N*INCX.  (Input/Output)
C              DSCAL replaces X(I) with DA*X(I) for I=1,...,N. X(I)
C              refers to a specific element of DX. See INCX argument
C              description.
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be DX(1+(I-1)*INCX). INCX must be
C              greater than zero.
C
C  GAMS:       D1a6
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
      SUBROUTINE DSCAL (N, DA, DX, INCX)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      DOUBLE PRECISION DA, DX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, M, MP1, NS
C
      IF (N .GT. 0) THEN
         IF (INCX .NE. 1) THEN
C                                  CODE FOR INCREMENTS NOT EQUAL TO 1.
            NS = N*INCX
            DO 10  I=1, NS, INCX
               DX(I) = DA*DX(I)
   10       CONTINUE
         ELSE
C                                  CODE FOR INCREMENTS EQUAL TO 1.
C                                  CLEAN-UP LOOP SO REMAINING VECTOR
C                                  LENGTH IS A MULTIPLE OF 5.
            M = N - (N/5)*5
            DO 30  I=1, M
               DX(I) = DA*DX(I)
   30       CONTINUE
            MP1 = M + 1
            DO 40  I=MP1, N, 5
               DX(I) = DA*DX(I)
               DX(I+1) = DA*DX(I+1)
               DX(I+2) = DA*DX(I+2)
               DX(I+3) = DA*DX(I+3)
               DX(I+4) = DA*DX(I+4)
   40       CONTINUE
         END IF
      END IF
      RETURN
      END
