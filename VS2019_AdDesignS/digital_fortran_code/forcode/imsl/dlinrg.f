C-----------------------------------------------------------------------
C  IMSL Name:  LINRG/DLINRG (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    February 27, 1985
C
C  Purpose:    Compute the inverse of a real general matrix.
C
C  Usage:      CALL LINRG (N, A, LDA, AINV, LDAINV)
C
C  Arguments:
C     N      - Order of the matrix A.  (Input)
C     A      - N by N matrix containing the matrix to be inverted.
C              (Input)
C     LDA    - Leading dimension of A exactly as specified in the
C              dimension statement of the calling program.  (Input)
C     AINV   - N by N matrix containing the inverse of A.  (Output)
C              If A is not needed, A and AINV can share the same storage
C              locations.
C     LDAINV - Leading dimension of AINV exactly as specified in the
C              dimension statement of the calling program.  (Input)
C
C  Remarks:
C  1. Automatic workspace usage is
C              LINRG     2*N + N*(N-1)/2 units, or
C              DLINRG    3*N + N*(N-1)   units.
C     Workspace may be explicitly provided, if desired, by use of
C     L2NRG/DL2NRG.  The reference is
C              CALL L2NRG (N, A, LDA, AINV, LDAINV, WK, IWK)
C     The additional arguments are as follows:
C     WK     - Work vector of length N + N*(N-1)/2.
C     IWK    - Integer work vector of length N.
C
C  2. Informational errors
C     Type Code
C       3   1  The input matrix is too ill-conditioned.  The inverse
C              might not be accurate.
C       4   2  The input matrix is singular.
C
C  Keywords:   Gaussian elimination; LU factorization
C
C  GAMS:       D2a1
C
C  Chapters:   MATH/LIBRARY Linear Systems
C              STAT/LIBRARY Mathematical Support
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C

      SUBROUTINE DLINRG (N, A, LDA, AINV, LDAINV,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, LDA, LDAINV,NFLAG,NFLAGBIS
      DOUBLE PRECISION A(LDA,*), AINV(LDAINV,*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    INDI, INDR
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
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, DL2NRG
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KGT, N1RCD
      INTEGER    I1KGT, N1RCD,II
C
      NFLAGBIS=0
      CALL E1PSH ('DLINRG ',NFLAG)
      IF(NFLAG.NE.0) GOTO 9999
C
      IF (N .LE. 0) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The order of the matrix must be '//
     &               'positive while N = %(I1) is given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF(NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=321
         GOTO 9000
      ELSE
         II=I1KGT(N,2,NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF(NFLAGBIS.NE.0) GOTO 9000
         INDI = II
         II=I1KGT(N+N*(N-1)/2,4,NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF(NFLAGBIS.NE.0) GOTO 9000
         INDR = II
         II=N1RCD(0,NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF(NFLAGBIS.NE.0) GOTO 9000
         IF (II.NE.0) THEN
            CALL E1MES (5, -1, ' ',NFLAG)
            NFLAGBIS=NFLAG
         NFLAG=0
         IF(NFLAGBIS.NE.0) GOTO 9000
            CALL E1STI (1, N)
            CALL E1MES (5, 2, 'The workspace is based on N, '//
     &                  'where N = %(I1).',NFLAG)
            NFLAGBIS=NFLAG
            NFLAG=0
            IF(NFLAGBIS.NE.0) GOTO 9000
            NFLAGBIS=322
            GOTO 9000
         ELSE
            CALL DL2NRG (N, A, LDA, AINV, LDAINV, RDWKSP(INDR),
     &                   IWKSP(INDI),NFLAG)
            NFLAGBIS=NFLAG
            NFLAG=0
            IF (NFLAGBIS.NE.0) GOTO 9000
         END IF
      END IF
C
 9000 CALL E1POP ('DLINRG ',NFLAG)
C      IF(NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
 9999 RETURN
      END
