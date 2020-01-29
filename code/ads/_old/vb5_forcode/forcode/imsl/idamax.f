C-----------------------------------------------------------------------
C  IMSL Name:  IDAMAX (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Find the smallest index of the component of a
C              double-precision vector having maximum absolute value.
C
C  Usage:      IDAMAX(N, DX, INCX)
C
C  Arguments:
C     N      - Length of vector X.  (Input)
C     DX     - Double precision vector of length N*INCX.  (Input)
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be DX(1+(I-1)*INCX). INCX must be
C              greater than zero.
C     IDAMAX - The smallest index I such that DABS(X(I)) is the maximum
C              of DABS(X(J)) for J=1 to N.  (Output)
C              X(I) refers to a specific element of DX. See INCX
C              argument description.
C
C  GAMS:       D1a2
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
      INTEGER FUNCTION IDAMAX (N, DX, INCX)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      DOUBLE PRECISION DX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, II, NS
      DOUBLE PRECISION DMAX, XMAG
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS
      INTRINSIC  DABS
      DOUBLE PRECISION DABS
C
      IDAMAX = 0
      IF (N .GE. 1) THEN
         IDAMAX = 1
         IF (N .GT. 1) THEN
            IF (INCX .NE. 1) THEN
C                                  CODE FOR INCREMENTS NOT EQUAL TO 1.
               DMAX = DABS(DX(1))
               NS = N*INCX
               II = 1
               DO 10  I=1, NS, INCX
                  XMAG = DABS(DX(I))
                  IF (XMAG .GT. DMAX) THEN
                     IDAMAX = II
                     DMAX = XMAG
                  END IF
                  II = II + 1
   10          CONTINUE
            ELSE
C                                  CODE FOR INCREMENTS EQUAL TO 1.
               DMAX = DABS(DX(1))
               DO 20  I=2, N
                  XMAG = DABS(DX(I))
                  IF (XMAG .GT. DMAX) THEN
                     IDAMAX = I
                     DMAX = XMAG
                  END IF
   20          CONTINUE
            END IF
         END IF
      END IF
      RETURN
      END
