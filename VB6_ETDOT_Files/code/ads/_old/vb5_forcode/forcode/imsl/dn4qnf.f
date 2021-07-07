C-----------------------------------------------------------------------
C  IMSL Name:  N4QNF/DN4QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    October 1, 1985
C
C  Purpose:    Calculate the Jacobian in order to solve a system of
C              nonlinear equations
C
C  Usage:      CALL N4QNF (FCN, N, X, FVEC, FJAC, IFLAG, ML, MU, EPSFCN,
C                          WK1, WK2)
C
C  Arguments:
C     FCN    - A real function subroutine supplied by the user.  FCN
C              must be declared EXTERNAL in the CALLING PROGRAM.
C              FCN specifies the system of equations to be solved and
C              should be of the following form
C                             SUBROUTINE FCN(X,F,N)
C                             REAL X(*),F(*)
C                             F(1)=
C                             .
C                             F(N)=
C                             RETURN
C                             END
C              Where X is given.  FCN must not alter X.
C     N      - The number of equations to be solbed and the number
C              of unknowns.  (Input)
C     X      - A vector of length N.  X contains the best estimate
C              of the root found by NEQNF.  (Output)
C     FVEC   - A vector of length N.  FVEC contains the functions
C              evaluated at the point X.  (Input)
C     FJAC   - An N by N matrix.  FJAC contains the orthogonal
C              matrix Q produced by the QR factorization of the
C              final approximate Jacobian.  (Output)
C     IFLAG  -
C     ML     -
C     MU     -
C     EPSFCN -
C     WK1    - Real work array of length N.  (Output)
C     WK2    - Real work array of length N.  (Output)
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN4QNF (FCN, N, X, FVEC, FJAC, IFLAG, ML, MU, EPSFCN,
     &                   WK1, WK2, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, IFLAG, ML, MU
      DOUBLE PRECISION EPSFCN, X(*), FVEC(*), FJAC(N,*), WK1(*), WK2(*)
      EXTERNAL   FCN
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, J, K, MSUM
      DOUBLE PRECISION EPS, EPSMCH, H, TEMP
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS,DMAX1,DSQRT
      INTRINSIC  DABS, DMAX1, DSQRT
      DOUBLE PRECISION DABS, DMAX1, DSQRT
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1USR
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH
      DOUBLE PRECISION DMACH
C
c      write(9,*) 'Enter dn4qnf.'


      EPSMCH = DMACH(4, nflag)
      EPS = DSQRT(DMAX1(EPSFCN,EPSMCH))
      MSUM = ML + MU + 1
      IF (MSUM .GE. N) THEN
C                                  Computation of dense approxiamte
C                                  JACOBIAN
         DO 20  J=1, N
            TEMP = X(J)
            H = EPS*DABS(TEMP)
            IF (H .EQ. 0.0D0) H = EPS
            X(J) = TEMP + H
            CALL E1USR ('ON', nflag)
            CALL FCN (X, WK1, N)
            CALL E1USR ('OFF', nflag)
            IF (IFLAG .LT. 0) GO TO 9000
            X(J) = TEMP
            DO 10  I=1, N
               FJAC(I,J) = (WK1(I)-FVEC(I))/H
   10       CONTINUE
   20    CONTINUE
         GO TO 9000
      END IF
C                                  Computation of banded approximate
C                                  JACOBIAN
      DO 60  K=1, MSUM
         DO 30  J=K, N, MSUM
            WK2(J) = X(J)
            H = EPS*DABS(WK2(J))
            IF (H .EQ. 0.0D0) H = EPS
            X(J) = WK2(J) + H
   30    CONTINUE
         CALL E1USR ('ON', nflag)
         CALL FCN (X, WK1, N)
         CALL E1USR ('OFF', nflag)
         IF (IFLAG .LT. 0) GO TO 9000
         DO 50  J=K, N, MSUM
            X(J) = WK2(J)
            H = EPS*DABS(WK2(J))
            IF (H .EQ. 0.0D0) H = EPS
            DO 40  I=1, N
               FJAC(I,J) = 0.0D0
               IF (I.GE.J-MU .AND. I.LE.J+ML) FJAC(I,J) =
     &             (WK1(I)-FVEC(I))/H
   40       CONTINUE
   50    CONTINUE
   60 CONTINUE

c      write(9,*) 'Exit dn4qnf.'

 9000 RETURN
      END
