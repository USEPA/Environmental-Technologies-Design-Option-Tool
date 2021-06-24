C-----------------------------------------------------------------------
C  IMSL Name:  N2QNF/DN2QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    July 16, 1985
C
C  Purpose:    Solve a system of nonlinear equations using the
C              Levenberg-Marquardt algorithm and a finite difference
C              Jacobian.
C
C  Usage:      CALL N2QNF (FCN, ERRREL, N, ITMAX, XINIT, X,
C                          FNORM, FVEC, FJAC, R, QTF, WK)
C
C  Arguments:  (See NEQNF)
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN2QNF (FCN, ERRREL, N, ITMAX, XINIT, X, FNORM, FVEC,
     &                   FJAC, R, QTF, WK, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, ITMAX
      DOUBLE PRECISION ERRREL, FNORM, XINIT(*), X(*), FVEC(*),
     &           FJAC(N,*), R(*), QTF(*), WK(*)
      EXTERNAL   FCN
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    INFO, ITER, LR, MAXFEV, ML, MODE, MU, NFEV, NPRINT
      DOUBLE PRECISION EPSFCN
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      DOUBLE PRECISION FACTOR
      SAVE       FACTOR
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, E1STD, DCOPY, DSET, DN3QNF
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   N1RTY, DDOT
      INTEGER    N1RTY
      DOUBLE PRECISION DDOT
C
      DATA FACTOR/1.0D2/
C
c      print *, 'Enter dn2qnf.'

      CALL E1PSH ('DN2QNF ', nflag)

c      print *, 'errrel = ', errrel
c      print *, 'n = ', n
c      print *, 'itmax = ', itmax

C
      IF (N .LT. 1) THEN
c         print *, 'dn2qnf: point a.'
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The argument N = %(I1).  The '//
     &               'number of equations to be solved and the '//
     &               'number of unknowns must be at least equal to '//
     &               '1.', nflag)
      END IF
C
      IF (ERRREL .LT. 0) THEN
c         print *, 'dn2qnf: point b.'
         CALL E1STD (1, ERRREL)
         CALL E1MES (5, 2, 'The argument ERRREL = %(D1).  The '//
     &               'bound for the relative error should be at '//
     &               'least equal to 0.', nflag)
      END IF
C
      IF (ITMAX .LT. 1) THEN
c         print *, 'dn2qnf: point c.'
         CALL E1STI (1, ITMAX)
         CALL E1MES (5, 3, 'The argument ITMAX = %(I1).  The '//
     &               'maximum number of iterations must be at '//
     &               'least 1.', nflag)
      END IF
c      print *, 'dn2qnf: point d.'
      IF (N1RTY(0, nflag) .NE. 0) GO TO 9000
C
      INFO = 0
      MAXFEV = ITMAX*(N+1)
      ML = N - 1
      MU = N - 1
      EPSFCN = 0.0D0
      MODE = 2
      CALL DSET (N, 1.0D0, WK, 1)
      NPRINT = 0
      LR = (N*(N+1))/2
c      print *, 'dn2qnf: point e.'
C                                  Copy initial guesses into X
      CALL DCOPY (N, XINIT, 1, X, 1)
C
c      print *, 'dn2qnf: point f.'
      CALL DN3QNF (FCN, ERRREL, N, X, FVEC, FJAC, R, QTF, MAXFEV, ML,
     &             MU, EPSFCN, MODE, FACTOR, NPRINT, INFO, NFEV, LR,
     &             WK(1), WK(N+1), WK(2*N+1), WK(3*N+1), WK(4*N+1),
     &             nflag)
C
      IF (INFO .EQ. 5) INFO = 4
      FNORM = DDOT(N,FVEC,1,FVEC,1)
C
c      print *, 'dn2qnf: point g.'
      IF (INFO .EQ. 2) THEN
         ITER = ITMAX*(N+1)
         CALL E1STI (1, ITER)
         CALL E1MES (4, 1, 'The number of calls to the function has '//
     &               'exceeded ITMAX*(N+1) = %(I1).  The user may '//
     &               'try a new initial guess.', nflag)
      ELSE IF (INFO .EQ. 3) THEN
         CALL E1STD (1, ERRREL)
         CALL E1MES (4, 2, 'The bound for the relative error, ERRREL '//
     &               '= %(D1), is too small.  No further improvement '//
     &               'in the approximate solution is possible.  The '//
     &               'user should increase ERRREL.', nflag)
      ELSE IF (INFO .EQ. 4) THEN
         CALL E1MES (4, 3, 'The iteration has not made good '//
     &               'progress.  The user may try a new initial '//
     &               'guess.', nflag)
      END IF
C
 9000 CALL E1POP ('DN2QNF ', nflag)
C
c      print *, 'Exit dn2qnf.'

      RETURN
      END
