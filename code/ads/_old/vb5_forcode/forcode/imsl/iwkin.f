C-----------------------------------------------------------------------
C  IMSL Name:  IWKIN (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    January 17, 1984
C
C  Purpose:    Initialize bookkeeping locations describing the
C              workspace stack.
C
C  Usage:      CALL IWKIN (NSU)
C
C  Argument:
C     NSU    - Number of numeric storage units to which the workspace
C              stack is to be initialized
C
C  GAMS:       N4
C
C  Chapters:   MATH/LIBRARY Reference Material
C              STAT/LIBRARY Reference Material
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C

      SUBROUTINE IWKIN (NSU,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NSU,NFLAG
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    ISIZE(6), LALC, LBND, LBOOK, LMAX, LNEED, LNOW, LOUT,
     &           LUSED, MELMTS, MTYPE
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    FIRST
      SAVE       FIRST
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                                  SPECIFICATIONS FOR COMMON /WORKSP/
      REAL       RWKSP(5000)
      REAL       RDWKSP(5000)
      DOUBLE PRECISION DWKSP(2500)
      COMPLEX    CWKSP(2500)
      COMPLEX    CZWKSP(2500)
      COMPLEX    *16 ZWKSP(1250)
      INTEGER    IWKSP(5000)
      LOGICAL    LWKSP(5000)
      EQUIVALENCE (DWKSP(1), RWKSP(1))
      EQUIVALENCE (CWKSP(1), RWKSP(1)), (ZWKSP(1), RWKSP(1))
      EQUIVALENCE (IWKSP(1), RWKSP(1)), (LWKSP(1), RWKSP(1))
      EQUIVALENCE (RDWKSP(1), RWKSP(1)), (CZWKSP(1), RWKSP(1))
      COMMON     /WORKSP/ RWKSP
C                                  SPECIFICATIONS FOR EQUIVALENCE
      EQUIVALENCE (LOUT, IWKSP(1))
      EQUIVALENCE (LNOW, IWKSP(2))
      EQUIVALENCE (LUSED, IWKSP(3))
      EQUIVALENCE (LBND, IWKSP(4))
      EQUIVALENCE (LMAX, IWKSP(5))
      EQUIVALENCE (LALC, IWKSP(6))
      EQUIVALENCE (LNEED, IWKSP(7))
      EQUIVALENCE (LBOOK, IWKSP(8))
      EQUIVALENCE (ISIZE(1), IWKSP(11))
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  MAX0
      INTRINSIC  MAX0
      INTEGER    MAX0
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1STI
C
      DATA FIRST/.TRUE./
C

      IF (.NOT.FIRST) THEN
         IF (NSU .NE. 0) THEN
            CALL E1STI (1, LMAX)
            CALL E1MES (5, 100, 'Error from subroutine IWKIN:  '//
     &                  'Workspace stack has previously been '//
     &                  'initialized to %(I1). Correct by making the '//
     &                  'call to IWKIN the first executable '//
     &                  'statement in the main program.  ',NFLAG)
          IF (NFLAG.NE.0) GOTO 1000
C
c            STOP
        NFLAG=97
        GOTO 1000
C
         ELSE
            RETURN
         END IF
      END IF
C
      IF (NSU .EQ. 0) THEN
C                                  IF NSU=0 USE DEFAULT SIZE 5000
         MELMTS = 5000
      ELSE
         MELMTS = NSU
      END IF
C                                  NUMBER OF ITEMS .LT. 0
      IF (MELMTS .LE. 0) THEN
         CALL E1STI (1, MELMTS)
         CALL E1MES (5, 1, 'Error from subroutine IWKIN:  Number '//
     &               'of numeric storage units is not positive. NSU '//
     &               '= %(I1) ',NFLAG)
        IF (NFLAG.NE.0) GOTO 1000
      ELSE
C
         FIRST = .FALSE.
C                                  HERE TO INITIALIZE
C
C                                  SET DATA SIZES APPROPRIATE FOR A
C                                  STANDARD CONFORMING FORTRAN SYSTEM
C                                  USING THE FORTRAN
C                                  *NUMERIC STORAGE UNIT* AS THE
C                                  MEASURE OF SIZE.
C
C                                  TYPE IS REAL
         MTYPE = 3
C                                  LOGICAL
         ISIZE(1) = 1
C                                  INTEGER
         ISIZE(2) = 1
C                                  REAL
         ISIZE(3) = 1
C                                  DOUBLE PRECISION
         ISIZE(4) = 2
C                                  COMPLEX
         ISIZE(5) = 2
C                                  DOUBLE COMPLEX
         ISIZE(6) = 4
C                                  NUMBER OF WORDS USED FOR BOOKKEEPING
         LBOOK = 16
C                                  CURRENT ACTIVE LENGTH OF THE STACK
         LNOW = LBOOK
C                                  MAXIMUM VALUE OF LNOW ACHIEVED THUS
C                                  FAR
         LUSED = LBOOK
C                                  MAXIMUM LENGTH OF THE STORAGE ARRAY
         LMAX = MAX0(MELMTS,((LBOOK+2)*ISIZE(2)+ISIZE(3)-1)/ISIZE(3))
C                                  LOWER BOUND OF THE PERMANENT STORAGE
C                                  WHICH IS ONE WORD MORE THAN THE
C                                  MAXIMUM ALLOWED LENGTH OF THE STACK
         LBND = LMAX + 1
C                                  NUMBER OF CURRENT ALLOCATIONS
         LOUT = 0
C                                  TOTAL NUMBER OF ALLOCATIONS MADE
         LALC = 0
C                                  NUMBER OF WORDS BY WHICH THE ARRAY
C                                  SIZE MUST BE INCREASED FOR ALL PAST
C                                  ALLOCATIONS TO SUCCEED
         LNEED = 0
      END IF
C
1000  RETURN
      END
