C-----------------------------------------------------------------------
C  IMSL Name:  N1RNOF
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    August 2, 1985
C
C  Purpose:    Set or retrieve the error checksum number-of-failures
C              flag.
C
C  Usage:      N1RNOF (IOPT)
C
C  Arguments:
C     IOPT   - Integer specifying the desired option.  (Input)
C              If IOPT=1 the number-of-failures flag is increased by 1.
C              If IOPT=2 the number-of-failures flag value is returned
C                        in N1RNOF and the flag is set to zero.
C     N1RNOF - Integer function. (Output)  The number-of-failures flag
C              value is returned in N1RNOF.
C
C  Copyright:  1985 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      INTEGER FUNCTION N1RNOF (IOPT,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IOPT,NFLAG
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    NOF
      SAVE       NOF
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                              SPECIFICATIONS FOR COMMON /ERCOM1/
      INTEGER    CALLVL, MAXLEV, MSGLEN, ERTYPE(51), ERCODE(51),
     &           PRINTB(7), STOPTB(7), PLEN, IFERR6, IFERR7,
     &           IALLOC(51), HDRFMT(7), TRACON(7)
      COMMON     /ERCOM1/ CALLVL, MAXLEV, MSGLEN, ERTYPE, ERCODE,
     &           PRINTB, STOPTB, PLEN, IFERR6, IFERR7, IALLOC, HDRFMT,
     &           TRACON
      SAVE       /ERCOM1/
C                              SPECIFICATIONS FOR COMMON /ERCOM2/
      CHARACTER  MSGSAV(255), PLIST(300), RNAME(51)*6
      COMMON     /ERCOM2/ MSGSAV, PLIST, RNAME
      SAVE       /ERCOM2/
C                              SPECIFICATIONS FOR COMMON /ERCOM3/
      DOUBLE PRECISION ERCKSM
      COMMON     /ERCOM3/ ERCKSM
      SAVE       /ERCOM3/
C                              SPECIFICATIONS FOR COMMON /ERCOM4/
      LOGICAL    ISUSER(51)
      COMMON     /ERCOM4/ ISUSER
      SAVE       /ERCOM4/
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1PRT, M1VECH
C
      DATA NOF/0/
C
      IF (IOPT .EQ. 1) THEN
C                                  INCREMENT NO.-OF-FAILURES FLAG
         NOF = NOF + 1
         N1RNOF = NOF
      ELSE IF (IOPT .EQ. 2) THEN
C                                  RETRIEVE NO.-OF-FAILURES FLAG
         N1RNOF = NOF
C                                  CLEAR NO.-OF-FAILURES FLAG
         NOF = 0
      ELSE
         ERTYPE(CALLVL) = 5
         ERCODE(CALLVL) = 1
         MSGLEN = 49
         CALL M1VECH ('.  The argument passed to N1RNOF must be 1 '//
     &                'or 2. ', MSGLEN, MSGSAV, MSGLEN)
	 
         CALL E1PRT(NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 NFLAG=149
C         STOP
      END IF
C
9999  RETURN
      END
