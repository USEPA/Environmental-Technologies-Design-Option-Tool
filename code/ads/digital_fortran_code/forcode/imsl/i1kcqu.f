C-----------------------------------------------------------------------
C  IMSL Name:  I1KCQU
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    January 19, 1984
C
C  Purpose:    Return number of CHARACTER*(LEN) locations that remain
C              to be allocated in one request.
C
C  Usage:      I1KCQU (LEN)
C
C  Arguments:
C     LEN    - Length of the character string to be checked. (Input)
C     I1KCQU - Integer function. (Output) Returns number of character
C              elements of length LEN remaining in the stack.
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C

      INTEGER FUNCTION I1KCQU (LEN,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    LEN,NFLAG,NFLAGBIS
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    LALC, LBND, LBOOK, LBPERM, LMAX, LMAXA, LNEED,
     &           LNEEDA, LNOW, LOUT, LUSED, NDXS
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    INIT
      SAVE       INIT
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
C                                  SPECIFICATIONS FOR COMMON /WKSPCH/
      CHARACTER  *1 CHWKSP(2000)
      COMMON     /WKSPCH/ CHWKSP
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  MAX0
      INTRINSIC  MAX0
      INTEGER    MAX0
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, I1KC00
C
      DATA INIT/.TRUE./
C
      NFLAGBIS=0
      CALL E1PSH ('I1KCQU',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (INIT) CALL I1KC00 (INIT, 2000, 20,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (LEN .LE. 0) THEN
	 CALL E1STI (1, LEN)
	 CALL E1MES (5, 11, 'Length of the character string for '//
     &               'which the check is to be made must be greater '//
     &               'than 0. LEN = %(I1).',NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 NFLAGBIS=150
      ELSE
	 NDXS = IWKSP(9)
C                                  CHECK POINTERS IN NUMERICAL
C                                  BOOKKEEPING SECTION
	 IF (NDXS.LT.IWKSP(4) .OR. NDXS.GE.IWKSP(5)) THEN
	    CALL E1MES (5, 12, 'Stack pointers have been '//
     &                  'overwritten in integer array IWKSP. ',NFLAG)
	    IF (NFLAG.NE.0) GOTO 9999
	    NFLAGBIS=151
	 ELSE
C
	    LOUT = IWKSP(NDXS)
	    LNOW = IWKSP(NDXS+1)
	    LUSED = IWKSP(NDXS+2)
	    LBND = IWKSP(NDXS+3)
	    LMAX = IWKSP(NDXS+4)
	    LALC = IWKSP(NDXS+5)
	    LNEED = IWKSP(NDXS+6)
	    LBOOK = IWKSP(NDXS+7)
	    LBPERM = IWKSP(NDXS+8)
	    LMAXA = IWKSP(NDXS+9)
	    LNEEDA = IWKSP(NDXS+10)
C                                  CHECK CHARACTER WORKSPACE POINTERS
C                                  WHICH RESIDE IN A PART OF THE
C                                  PERMANENT SECTION OF THE NUMERICAL
C                                  WORKSPACE ARRAY
	    IF (LNOW.LT.0 .OR. LNOW.GT.LUSED .OR. LUSED.GT.LMAX .OR.
     &          LNOW.GE.LBND .OR. LOUT.GT.LALC) THEN
	       CALL E1MES (5, 13, 'One or more of the character '//
     &                     'workspace bookkeeping locations have '//
     &                     'been overwritten.',NFLAG)
		IF (NFLAG.NE.0) GOTO 9999
		NFLAGBIS=152
	    ELSE
	       I1KCQU = MAX0((LBND-2)/LEN-(LNOW-1)/LEN-1,0)
	    END IF
	 END IF
      END IF
C
      CALL E1POP ('I1KCQU',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
9999  RETURN
      END
