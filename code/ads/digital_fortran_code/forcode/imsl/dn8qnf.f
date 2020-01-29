C-----------------------------------------------------------------------
C  IMSL Name:  N8QNF/DN8QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    October 1, 1985
C
C  Purpose:
C
C  Usage:      CALL N8QNF (M, N, S, U, V, W, SING)
C
C  Arguments:
C     M      -
C     N      -
C     S      -
C     U      -
C     V      -
C     W      -
C     SING   -
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN8QNF (M, N, S, U, V, W, SING, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    M, N
      DOUBLE PRECISION S(*), U(*), V(*), W(*)
      LOGICAL    SING
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, J, JJ, L, NM1, NMJ
      DOUBLE PRECISION GIANT, TAU, TEMP, TEMP1, TEMP2, TEMP3, TEMP4
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS,DSQRT
      INTRINSIC  DABS, DSQRT
      DOUBLE PRECISION DABS, DSQRT
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   DCOPY
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH
      DOUBLE PRECISION DMACH
C
      GIANT = DMACH(2, nflag)
      JJ = (N*(2*M-N+1))/2 - (M-N)
C                                  Move the nontrivial part of the last
C                                  column of S into W
      CALL DCOPY (M-N+1, S(JJ), 1, W(N), 1)
C                                  Rotate the vector V into a multiple
C                                  of the N-th unit vector in such a way
C                                  that a spike is introduced into W
      NM1 = N - 1
      IF (NM1 .GE. 1) THEN
         DO 20  NMJ=1, NM1
            J = N - NMJ
            JJ = JJ - (M-J+1)
            W(J) = 0.0D0
            IF (V(J) .NE. 0.0D0) THEN
C                                  Determine a GIVENS rotation which
C                                  eliminates the J-th element of V
               IF (DABS(V(N)) .LT. DABS(V(J))) THEN
                  TEMP2 = V(N)/V(J)
                  TEMP3 = 0.5D0/DSQRT(0.25D0+0.25D0*TEMP2**2)
                  TEMP1 = TEMP3*TEMP2
                  TAU = 1.0D0
                  IF (DABS(TEMP1)*GIANT .GT. 1.0D0) TAU = 1.0D0/TEMP1
               ELSE
                  TEMP4 = V(J)/V(N)
                  TEMP1 = 0.5D0/DSQRT(0.25D0+0.25D0*TEMP4**2)
                  TEMP3 = TEMP1*TEMP4
                  TAU = TEMP3
               END IF
C                                  Apply the transformation to V and
C                                  store the information necessary to
C                                  recover the GIVENS rotation
               V(N) = TEMP3*V(J) + TEMP1*V(N)
               V(J) = TAU
C                                  Apply the transformation to S and
C                                  extend the spike in W
               L = JJ
               DO 10  I=J, M
                  TEMP = TEMP1*S(L) - TEMP3*W(I)
                  W(I) = TEMP3*S(L) + TEMP1*W(I)
                  S(L) = TEMP
                  L = L + 1
   10          CONTINUE
            END IF
   20    CONTINUE
      END IF
C                                  Add the spike from the rank 1 update
C                                  to W
      DO 30  I=1, M
         W(I) = W(I) + V(N)*U(I)
   30 CONTINUE
C                                  Eliminate the spike
      SING = .FALSE.
      IF (NM1 .GE. 1) THEN
         DO 50  J=1, NM1
            IF (W(J) .NE. 0.0D0) THEN
C                                  Determine a GIVENS rotation which
C                                  eliminates the J-th element of the
C                                  spike
               IF (DABS(S(JJ)) .LT. DABS(W(J))) THEN
                  TEMP2 = S(JJ)/W(J)
                  TEMP3 = 0.5D0/DSQRT(0.25D0+0.25D0*TEMP2**2)
                  TEMP1 = TEMP3*TEMP2
                  TAU = 1.0D0
                  IF (DABS(TEMP1)*GIANT .GT. 1.0D0) TAU = 1.0D0/TEMP1
               ELSE
                  TEMP4 = W(J)/S(JJ)
                  TEMP1 = 0.5D0/DSQRT(0.25D0+0.25D0*TEMP4**2)
                  TEMP3 = TEMP1*TEMP4
                  TAU = TEMP3
               END IF
C                                  Apply the transformation to S and
C                                  reduce the spike in W
               L = JJ
               DO 40  I=J, M
                  TEMP = TEMP1*S(L) + TEMP3*W(I)
                  W(I) = -TEMP3*S(L) + TEMP1*W(I)
                  S(L) = TEMP
                  L = L + 1
   40          CONTINUE
C                                  Sotre the information necessary to
C                                  recover the GIVENS rotation
               W(J) = TAU
            END IF
C                                  Test for zero diagonal elements in
C                                  the output S
            IF (S(JJ) .EQ. 0.0D0) SING = .TRUE.
            JJ = JJ + (M-J+1)
   50    CONTINUE
      END IF
C                                  Move W back into the last column of
C                                  the output S
      CALL DCOPY (M-N+1, W(N), 1, S(JJ), 1)
      IF (S(JJ) .EQ. 0.0D0) SING = .TRUE.
      RETURN
      END
