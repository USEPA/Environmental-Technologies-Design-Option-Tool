C-----------------------------------------------------------------------
C  IMSL Name:  IWKCIN (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    January 19, 1984
C
C  Purpose:    Initialize bookkeeping locations describing the
C              character workspace stack.
C
C  Usage:      CALL IWKCIN (NELMTS, NALC)
C
C  Arguments:
C     NELMTS - Number of CHARACTER*1 locations that are to be available
C              in the stack.  (Input)
C     NALC   - Maximum number of active allocations available.  (Input)
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
      SUBROUTINE IWKCIN (NELMTS, NALC,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NELMTS, NALC,NFLAG,NFLAGBIS
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    INIT
      SAVE       INIT
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, I1KC00
C
      DATA INIT/.TRUE./
C
      NFLAGBIS=0
      CALL E1PSH ('IWKCIN',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (NELMTS .LE. 0) THEN
         CALL E1STI (1, NELMTS)
         CALL E1MES (5, 6, 'The maximum number of items that '//
     &               'are to be available in the character stack '//
     &               'must be greater than 0.  NELMTS = %(I1).',NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
	NFLAGBIS=147
      END IF
C
      IF (NALC .LE. 0) THEN
         CALL E1STI (1, NALC)
         CALL E1MES (5, 7, 'The maximum number of active '//
     &               'allocations allowed must be greater than 0. '//
     &               'NALC = %(I1).',NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
	NFLAGBIS=147
      END IF
C
      IF (INIT) CALL I1KC00 (INIT, NELMTS, NALC,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      CALL E1POP ('IWKCIN',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
9999  RETURN
      END
