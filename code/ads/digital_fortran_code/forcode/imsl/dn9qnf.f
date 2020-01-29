C-----------------------------------------------------------------------
C  IMSL Name:  N9QNF/DN9QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    October 1, 1985
C
C  Purpose:
C
C  Usage:      CALL N9QNF (M, N, A, V, W)
C
C  Arguments:
C     M      -
C     N      - The number of equations to be solbed and the number
C              of unknowns.  (Input)
C     A      -
C     V      -
C     W      -
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN9QNF (M, N, A, LDA, V, W, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    M, N, LDA
      DOUBLE PRECISION A(LDA,*), V(*), W(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, J, NM1, NMJ
      DOUBLE PRECISION TEMP, TEMP1, TEMP2
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS,DSQRT
      INTRINSIC  DABS, DSQRT
      DOUBLE PRECISION DABS, DSQRT
C
      NM1 = N - 1
      IF (NM1 .GE. 1) THEN
         DO 20  NMJ=1, NM1
            J = N - NMJ
            IF (DABS(V(J)) .GT. 1.0D0) THEN
               TEMP1 = 1.0D0/V(J)
               TEMP2 = DSQRT(1.0D0-TEMP1**2)
            ELSE
               TEMP2 = V(J)
               TEMP1 = DSQRT(1.0D0-TEMP2**2)
            END IF
            DO 10  I=1, M
               TEMP = TEMP1*A(I,J) - TEMP2*A(I,N)
               A(I,N) = TEMP2*A(I,J) + TEMP1*A(I,N)
               A(I,J) = TEMP
   10       CONTINUE
   20    CONTINUE
C                                  Apply the second set of GIVENS
C                                  rotations to A
         DO 40  J=1, NM1
            IF (DABS(W(J)) .GT. 1.0D0) THEN
               TEMP1 = 1.0D0/W(J)
               TEMP2 = DSQRT(1.0D0-TEMP1**2)
            ELSE
               TEMP2 = W(J)
               TEMP1 = DSQRT(1.0D0-TEMP2**2)
            END IF
            DO 30  I=1, M
               TEMP = TEMP1*A(I,J) + TEMP2*A(I,N)
               A(I,N) = -TEMP2*A(I,J) + TEMP1*A(I,N)
               A(I,J) = TEMP
   30       CONTINUE
   40    CONTINUE
      END IF
      RETURN
      END
