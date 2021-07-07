C-----------------------------------------------------------------------
C  IMSL Name:  DSET (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Set the components of a vector to a scalar, all double
C              precision.
C
C  Usage:      CALL DSET (N, DA, DX, INCX)
C
C  Arguments:
C     N      - Length of vector X.  (Input)
C     DA     - Double precision scalar.  (Input)
C     DX     - Double precison vector of length N*INCX.  (Input/Output)
C              DSET replaces X(I) with DA for I=1,...,N.  X(I) refers to
C              a specific element of DX. See INCX argument description.
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be DX(1+(I-1)*INCX).  INCX must be
C              greater than zero.
C
C  GAMS:       D1a1
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
      SUBROUTINE DSET (N, DA, DX, INCX)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      DOUBLE PRECISION DA, DX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, M, MP1, NINCX
C                                  SPECIFICATIONS FOR SPECIAL CASES
C     INTRINSIC  MOD
      INTRINSIC  MOD
      INTEGER    MOD
C
      IF (N .GT. 0) THEN
         IF (INCX .NE. 1) THEN
C                                  CODE FOR INCREMENT NOT EQUAL TO 1
            NINCX = N*INCX
            DO 10  I=1, NINCX, INCX
               DX(I) = DA
   10       CONTINUE
         ELSE
C                                  CODE FOR INCREMENT EQUAL TO 1
            M = MOD(N,8)
C                                  CLEAN-UP LOOP
            DO 30  I=1, M
               DX(I) = DA
   30       CONTINUE
            MP1 = M + 1
            DO 40  I=MP1, N, 8
               DX(I) = DA
               DX(I+1) = DA
               DX(I+2) = DA
               DX(I+3) = DA
               DX(I+4) = DA
               DX(I+5) = DA
               DX(I+6) = DA
               DX(I+7) = DA
   40       CONTINUE
         END IF
      END IF
      RETURN
      END
