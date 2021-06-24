C-----------------------------------------------------------------------
C  IMSL Name:  DCOPY (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Copy a vector X to a vector Y, both double precision.
C
C  Usage:      CALL DCOPY (N, DX, INCX, DY, INCY)
C
C  Arguments:
C     N      - Length of vectors X and Y.  (Input)
C     DX     - Double precision vector of length MAX(N*IABS(INCX),1).
C              (Input)
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be.. DX(1+(I-1)*INCX) if INCX .GE. 0
C              or DX(1+(I-N)*INCX) if INCX .LT. 0.
C     DY     - Double precision vector of length MAX(N*IABS(INCY),1).
C              (Output)
C              DCOPY copies X(I) to Y(I) for I=1,...,N. X(I) and Y(I)
C              refer to specific elements of DX and DY, respectively.
C              See INCX and INCY argument descriptions.
C     INCY   - Displacement between elements of DY.  (Input)
C              Y(I) is defined to be.. DY(1+(I-1)*INCY) if INCY .GE. 0
C              or DY(1+(I-N)*INCY) if INCY .LT. 0.
C
C  GAMS:       D1a
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
      SUBROUTINE DCOPY (N, DX, INCX, DY, INCY)
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
      IF (N .GT. 0) THEN
         IF (INCX.NE.1 .OR. INCY.NE.1) THEN
C                                  CODE FOR UNEQUAL INCREMENTS.
            IX = 1
            IY = 1
            IF (INCX .LT. 0) IX = (-N+1)*INCX + 1
            IF (INCY .LT. 0) IY = (-N+1)*INCY + 1
            DO 10  I=1, N
               DY(IY) = DX(IX)
               IX = IX + INCX
               IY = IY + INCY
   10       CONTINUE
         ELSE
C                                  CODE FOR BOTH INCREMENTS EQUAL TO 1
C                                  CLEAN-UP LOOP SO REMAINING VECTOR
C                                  LENGTH IS A MULTIPLE OF 7.
            M = MOD(N,7)
            DO 30  I=1, M
               DY(I) = DX(I)
   30       CONTINUE
            MP1 = M + 1
            DO 40  I=MP1, N, 7
               DY(I) = DX(I)
               DY(I+1) = DX(I+1)
               DY(I+2) = DX(I+2)
               DY(I+3) = DX(I+3)
               DY(I+4) = DX(I+4)
               DY(I+5) = DX(I+5)
               DY(I+6) = DX(I+6)
   40       CONTINUE
         END IF
      END IF
      RETURN
      END
