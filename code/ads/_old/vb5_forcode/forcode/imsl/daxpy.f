C-----------------------------------------------------------------------
C  IMSL Name:  DAXPY (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Compute the scalar times a vector plus a vector,
C              y = ax + y, all double precision.
C
C  Usage:      CALL DAXPY (N, DA, DX, INCX, DY, INCY)
C
C  Arguments:
C     N      - Length of vectors X and Y.  (Input)
C     DA     - Double precision scalar.  (Input)
C     DX     - Double precision vector of length MAX(N*IABS(INCX),1).
C                 (Input)
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be
C                 DX(1+(I-1)*INCX) if INCX.GE.0  or
C                 DX(1+(I-N)*INCX) if INCX.LT.0.
C     DY     - Double precision vector of length MAX(N*IABS(INCY),1).
C                 (Input/Output)
C              DAXPY replaces Y(I) with DA*X(I) + Y(I) for I=1,...,N.
C              X(I) and Y(I) refer to specific elements of DX and DY.
C     INCY   - Displacement between elements of DY.  (Input)
C              Y(I) is defined to be
C                 DY(1+(I-1)*INCY) if INCY.GE.0  or
C                 DY(1+(I-N)*INCY) if INCY.LT.0.
C
C  GAMS:       D1a7
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
      SUBROUTINE DAXPY (N, DA, DX, INCX, DY, INCY)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX, INCY
      DOUBLE PRECISION DA, DX(*), DY(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IX, IY, M, MP1
C                                  SPECIFICATIONS FOR SPECIAL CASES
C     INTRINSIC  MOD
      INTRINSIC  MOD
      INTEGER    MOD
C
      IF (N .GT. 0) THEN
         IF (DA .NE. 0.0D0) THEN
            IF (INCX.NE.1 .OR. INCY.NE.1) THEN
C                                  CODE FOR UNEQUAL INCREMENTS OR EQUAL
C                                  INCREMENTS NOT EQUAL TO 1
               IX = 1
               IY = 1
               IF (INCX .LT. 0) IX = (-N+1)*INCX + 1
               IF (INCY .LT. 0) IY = (-N+1)*INCY + 1
               DO 10  I=1, N
                  DY(IY) = DY(IY) + DA*DX(IX)
                  IX = IX + INCX
                  IY = IY + INCY
   10          CONTINUE
            ELSE
C                                  CODE FOR BOTH INCREMENTS EQUAL TO 1
               M = MOD(N,4)
C                                  CLEAN-UP LOOP
               DO 30  I=1, M
                  DY(I) = DY(I) + DA*DX(I)
   30          CONTINUE
               MP1 = M + 1
               DO 40  I=MP1, N, 4
                  DY(I) = DY(I) + DA*DX(I)
                  DY(I+1) = DY(I+1) + DA*DX(I+1)
                  DY(I+2) = DY(I+2) + DA*DX(I+2)
                  DY(I+3) = DY(I+3) + DA*DX(I+3)
   40          CONTINUE
            END IF
         END IF
      END IF
      RETURN
      END
