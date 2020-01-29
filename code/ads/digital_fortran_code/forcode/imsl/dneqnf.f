C-----------------------------------------------------------------------
C  IMSL Name:  NEQNF/DNEQNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    July 16, 1985
C
C  Purpose:    Solve a system of nonlinear equations using the
C              Levenberg-Marquardt algorithm and a finite-difference
C              Jacobian.
C
C  Usage:      CALL NEQNF (FCN, ERRREL, N, ITMAX, XINIT, X, FNORM)
C
C  Arguments:
C     FCN    - User-supplied SUBROUTINE to evaluate the system of
C              equations to be solved.  The usage is
C              CALL FCN (X, F, N), where
C              X      - The point at which the functions are evaluated.
C                       (Input)
C                       X should not be changed by FCN.
C              F      - The computed function values at the point X.
C                       (Output)
C              N      - Length of X and F.  (Input)
C              FCN must be declared EXTERNAL in the calling program.
C     ERRREL - Stopping criterion.  (Input)
C              The root is accepted if the relative error between two
C              successive approximations to this root is less than
C              ERRREL.
C     N      - The number of equations to be solved and the number
C              of unknowns.  (Input)
C     ITMAX  - The maximum allowable number of iterations.  (Input)
C              The maximum number of calls to FCN is ITMAX*(N+1).
C              Suggested value = 200.
C     XINIT  - A vector of length N.  (Input)
C              XINIT contains the initial estimate of the root.
C     X      - A vector of length N.  (Output)
C              X contains the best estimate of the root found by NEQNF.
C     FNORM  - A scalar which has the following value,
C              F(1)**2+...+F(N)**2 at the point X.  (Output)
C
C  Remarks:
C  1. Automatic workspace usage is
C              NEQNF     1.5*N**2 + 7.5*N   units, or
C              DNEQNF    3*N**2 + 15*N      units.
C     Workspace may be explicitly provided, if desired, by use of
C     N2QNF/DN2QNF.  The reference is
C              CALL N2QNF (FCN, ERRREL, N, ITMAX, XINIT, X, FNORM,
C                          FVEC, FJAC, R, QTF, WK)
C     The additional arguments are as follows:
C     FVEC   - A vector of length N.  FVEC contains the functions
C              evaluated at the point X.
C     FJAC   - An N by N matrix.  FJAC contains the orthogonal
C              matrix Q produced by the QR factorization of the
C              final approximate Jacobian.
C     R      - A vector of length N*(N+1)/2.  R contains the upper
C              triangular matrix produced by the QR factorization
C              of the final approximation Jacobian.  R is stored
C              row-wise.
C     QTF    - A vector of length N.  QTF contains the vector
C              TRANS(Q)*FVEC.
C     WK     - A work vector of length 5*N.
C
C  2. Informational errors
C     Type Code
C       4   1  The number of calls to FCN has exceeded ITMAX*(N+1).
C              A new initial guess may be tried.
C       3   2  ERRREL is too small.  No further improvement in the
C              approximate solution is possible.
C       4   3  The iteration has not made good progress.  A new
C              initial guess may be tried.
C
C  Keywords:   Powell hybrid method; Forward-difference approximation;
C              Roots
C
C  GAMS:       F2
C
C  Chapter:    MATH/LIBRARY Nonlinear Equations
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DNEQNF (FCN, ERRREL, N, ITMAX, XINIT, X, FNORM, NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, ITMAX
      DOUBLE PRECISION ERRREL, FNORM, XINIT(*), X(*)
      EXTERNAL   FCN
      integer nflag
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IFJAC, IFVEC, IQTF, IR, IWK, MXEVAL
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                                  SPECIFICATIONS FOR COMMON /WORKSP/
      REAL       RWKSP(5000)
      DOUBLE PRECISION RDWKSP(2500)
      DOUBLE PRECISION DWKSP(2500)
      COMPLEX    CWKSP(2500)
      COMPLEX    *16 CZWKSP(1250)
      COMPLEX    *16 ZWKSP(1250)
      INTEGER    IWKSP(5000)
      LOGICAL    LWKSP(5000)
      EQUIVALENCE (DWKSP(1), RWKSP(1))
      EQUIVALENCE (CWKSP(1), RWKSP(1)), (ZWKSP(1), RWKSP(1))
      EQUIVALENCE (IWKSP(1), RWKSP(1)), (LWKSP(1), RWKSP(1))
      EQUIVALENCE (RDWKSP(1), RWKSP(1)), (CZWKSP(1), RWKSP(1))
      COMMON     /WORKSP/ RWKSP
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, DN2QNF
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KGT, N1RTY
      INTEGER    I1KGT, N1RTY
C

      CALL E1PSH ('DNEQNF ', NFLAG)

c      print *, 'Enter dneqnf.'
c      print *, 'errrel = ', errrel
c      print *, 'n = ', n
c      print *, 'itmax = ', itmax

C                                  Check N
      IF (N .LT. 1) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The argument N = %(I1).  The '//
     &               'number of equations to be solved and the '//
     &               'number of unknowns must be at least equal to '//
     &               '1.', NFLAG)
c         print *, 'dneqnf point 10.'
      END IF
C                                  Check for errors
      IF (N1RTY(0, nflag) .NE. 0) then
c         print *, 'dneqnf point 11.'
         GO TO 9000
      end if
C                                  Allocate workspace
c      print *, 'dneqnf point a.'

      IFVEC = I1KGT(N,4, nflag)
      IFJAC = I1KGT(N*N,4, nflag)
      IR = I1KGT(N*(N+1)/2,4, nflag)
      IQTF = I1KGT(N,4, nflag)
      IWK = I1KGT(5*N,4, nflag)
C                                  Check for errors
      IF (N1RTY(0, nflag) .NE. 0) THEN
c         print *, 'dneqnf point b.'
         CALL E1MES (5, -1, ' ', NFLAG)
         CALL E1STI (1, N)
         CALL E1STI (2, MXEVAL)
         CALL E1MES (5, 2, 'The workspace requirement is based on '//
     &               'N = %(I1).', NFLAG)
         GO TO 9000
      END IF
C
c      print *, 'dneqnf point c.'
      CALL DN2QNF (FCN, ERRREL, N, ITMAX, XINIT, X, FNORM,
     &             RDWKSP(IFVEC), RDWKSP(IFJAC), RDWKSP(IR),
     &             RDWKSP(IQTF), RDWKSP(IWK), nflag)

c      SUBROUTINE DN2QNF (FCN, ERRREL, N, ITMAX, XINIT, X, FNORM, FVEC,
c     &                   FJAC, R, QTF, WK)

C
 9000 CALL E1POP ('DNEQNF ', nflag)

c      print *, 'Exit dneqnf.'

      RETURN
      END
