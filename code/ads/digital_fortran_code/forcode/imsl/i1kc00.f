C-----------------------------------------------------------------------
C  IMSL Name:  I1KC00
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    February 13, 1984
C
C  Purpose:    Nucleus called by character workspace allocation routine
C              IWKCIN to initialize the character stack to NELMTS
C              character strings and allow up to NALC active
C              allocations.
C
C  Usage:      CALL I1KC00 (LARG, NELMTS, NALC)
C
C  Arguments:
C     LARG   - Logical argument used to determine if I1KC00 has
C              previously been called. (Input/Output)
C     NELMTS - Number of CHARACTER*1 elements to which the stack is to
C              be initialized. (Input)
C     NALC   - Maximum number of active (unreleased) allocations.(Input)
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE I1KC00 (LARG, NELMTS, NALC,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NELMTS, NALC,NFLAG,NFLAGBIS,II
      LOGICAL    LARG
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    LBPERM, LMAX, NDXS
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    LALC, LBOOK, LMAXA, LNEED, LNEEDA, LNOW, LOUT, LUSED
      LOGICAL    INIT
      SAVE       INIT, LALC, LBOOK, LMAXA, LNEED, LNEEDA, LNOW, LOUT,
     &           LUSED
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
C     INTRINSIC  IABS
      INTRINSIC  IABS
      INTEGER    IABS
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PRT, E1PSH, E1STI
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KGT
      INTEGER    I1KGT
C
      DATA LOUT/0/, LNOW/0/, LUSED/0/, LALC/0/, LNEED/0/, LBOOK/11/
      DATA LMAXA/0/, LNEEDA/0/
      DATA INIT/.FALSE./
C
      NFLAGBIS=0
      LARG = .FALSE.
      IF (INIT) RETURN
C
      INIT = .TRUE.
C
      LBPERM = LBOOK + 2*NALC
      II=I1KGT(LBPERM,-2,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NDXS =II 
C                                  DETERMINE IF CHARACTER ALLOCATIONS
C                                  REQUIRE MORE SPACE THAN IS AVAILABLE
C                                  IN IWKSP
      IF (NDXS .LE. 0) THEN
	 CALL E1PSH ('I1KC00',NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 CALL E1STI (1, IWKSP(2)+IABS(NDXS))
	 CALL E1MES (5, 25, 'An attempt was made to initialize '//
     &               'character workspace so large that there is '//
     &               'insufficient space available in the integer '//
     &               'stack to hold pointers. A call to IWKIN '//
     &               'in the main program requesting AT LEAST %(I1) '//
     &               'storage units MAY be adequate.',NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 NFLAGBIS=153
C                                  PRINT THE MESSAGE EVEN IF THE ERROR
C                                  HANDLER PRINT ATTRIBUTE IS OFF
	 CALL E1PRT(NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 CALL E1POP ('I1KC00',NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 NFLAG=NFLAGBIS
	 GOTO 9999
C         STOP
      END IF
C
      IWKSP(9) = NDXS
      IWKSP(10) = LBPERM
C
      IWKSP(NDXS) = LOUT
      IWKSP(NDXS+1) = LNOW
      IWKSP(NDXS+2) = LUSED
      LMAX = NELMTS
      IWKSP(NDXS+3) = LMAX + 1
      IWKSP(NDXS+4) = LMAX
      IWKSP(NDXS+5) = LALC
      IWKSP(NDXS+6) = LNEED
      IWKSP(NDXS+7) = LBOOK
      IWKSP(NDXS+8) = LBPERM
      IWKSP(NDXS+9) = NALC
      IWKSP(NDXS+10) = LNEEDA
C
9999  RETURN
      END
