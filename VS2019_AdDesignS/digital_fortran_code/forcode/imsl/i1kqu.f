C-----------------------------------------------------------------------
C  IMSL Name:  I1KQU
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    January 17, 1984
C
C  Purpose:    Return number of elements of data type ITYPE that
C              remain to be allocated in one request.
C
C  Usage:      I1KQU(ITYPE)
C
C  Arguments:
C     ITYPE  - Type of storage to be checked (Input)
C                 1 - logical
C                 2 - integer
C                 3 - real
C                 4 - double precision
C                 5 - complex
C                 6 - double complex
C     I1KQU  - Integer function. (Output) Returns number of elements
C              of data type ITYPE remaining in the stack.
C
C  Copyright:  1983 by IMSL, Inc.  All Rights Reserved
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C

      INTEGER FUNCTION I1KQU (ITYPE,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    ITYPE,NFLAG,NFLAGBIS
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    ISIZE(6), LALC, LBND, LBOOK, LMAX, LNEED, LNOW, LOUT,
     &           LUSED
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
      EXTERNAL   E1MES, E1POP, E1PSH, IWKIN
C
      DATA FIRST/.TRUE./
C
      NFLAGBIS=0   
      CALL E1PSH ('I1KQU ',NFLAG)
      IF (NFLAG.NE.0) GOTO 1000
C
      IF (FIRST) THEN
C                                  INITIALIZE WORKSPACE IF NEEDED
         FIRST = .FALSE.
         CALL IWKIN (0,NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
      END IF
C                                  BOOKKEEPING OVERWRITTEN
      IF (LNOW.LT.LBOOK .OR. LNOW.GT.LUSED .OR. LUSED.GT.LMAX .OR.
     &    LNOW.GE.LBND .OR. LOUT.GT.LALC) THEN
         CALL E1MES (5, 7, 'One or more of the first eight '//
     &               'bookkeeping locations in IWKSP have been '//
     &               'overwritten.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=1
         GOTO 9000
      ELSE IF (ITYPE.LE.0 .OR. ITYPE.GE.7) THEN
C                                  ILLEGAL DATA TYPE REQUESTED
         CALL E1MES (5, 8, 'Illegal data type requested.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=2
         GOTO 9000
      ELSE
C                                  THIS CALCULATION ALLOWS FOR THE
C                                  TWO POINTER LOCATIONS IN THE STACK
C                                  WHICH ARE ASSIGNED TO EACH ALLOCATION
         I1KQU = MAX0(((LBND-3)*ISIZE(2))/ISIZE(ITYPE)-(LNOW*ISIZE(2)-
     &           1)/ISIZE(ITYPE)-1,0)
      END IF
C
9000  CALL E1POP ('I1KQU ',NFLAG)
      IF (NFLAG.NE.0) GOTO 1000
      NFLAG=NFLAGBIS
1000  RETURN
      END
