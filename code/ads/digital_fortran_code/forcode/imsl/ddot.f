C-----------------------------------------------------------------------
C  IMSL Name:  DDOT (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Compute the double-precision dot product x*y.
C
C  Usage:      DDOT(N, DX, INCX, DY, INCY)
C
C  Arguments:
C     N      - Length of vectors X and Y.  (Input)
C     DX     - Double precision vector of length MAX(N*IABS(INCX),1).
C              (Input)
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be.. DX(1+(I-1)*INCX) if INCX .GE. 0
C              or DX(1+(i-n)*INCX) if INCX .LT. 0.
C     DY     - Double precision vector of length MAX(N*IABS(INCY),1).
C              (Input)
C     INCY   - Displacement between elements of DY.  (Input)
C              Y(I) is defined to be.. DY(1+(I-1)*INCY) if INCY .GE. 0
C              or DY(1+(I-N)*INCY) if INCY .LT. 0.
C     DDOT   - Double precision sum from I=1 to N of X(I)*Y(I).
C              (Output)
C              X(I) and Y(I) refer to specific elements of DX and DY,
C              respectively. See INCX and INCY argument descriptions.
C
C  GAMS:       D1a4
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
      DOUBLE PRECISION FUNCTION DDOT (N, DX, INCX, DY, INCY)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX, INCY
      DOUBLE PRECISION DX(*), DY(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IX, IY, M, MP1
C                                  SPECIFICATIONS FOR SPECIAL CASES
C     INTRINSIC  MOD
      INTRINSIC  MOD
      INTEGER    MOD
C
      DDOT = 0.0D0
      IF (N .GT. 0) THEN
         IF (INCX.NE.1 .OR. INCY.NE.1) THEN
C                                  CODE FOR UNEQUAL INCREMENTS.
            IX = 1
            IY = 1
            IF (INCX .LT. 0) IX = (-N+1)*INCX + 1
            IF (INCY .LT. 0) IY = (-N+1)*INCY + 1
            DO 10  I=1, N
               DDOT = DDOT + DX(IX)*DY(IY)
               IX = IX + INCX
               IY = IY + INCY
   10       CONTINUE
         ELSE
C                                  CODE FOR BOTH INCREMENTS EQUAL TO 1
C                                    CLEAN-UP LOOP SO REMAINING VECTOR
C                                    LENGTH IS A MULTIPLE OF 5.
            M = MOD(N,5)
            DO 30  I=1, M
               DDOT = DDOT + DX(I)*DY(I)
   30       CONTINUE
            MP1 = M + 1
            DO 40  I=MP1, N, 5
               DDOT = DDOT + DX(I)*DY(I) + DX(I+1)*DY(I+1) +
     &                DX(I+2)*DY(I+2) + DX(I+3)*DY(I+3) +
     &                DX(I+4)*DY(I+4)
   40       CONTINUE
         END IF
      END IF
      RETURN
      END
