C-----------------------------------------------------------------------
C  IMSL Name:  N1RTY
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 6, 1984
C
C  Purpose:    Retrieve an error type.
C
C  Usage:      N1RTY(IOPT)
C
C  Arguments:
C     IOPT   - Integer specifying the level.  (Input)
C              If IOPT=0 the error type for the current level is
C              returned.  If IOPT=1 the error type for the most
C              recently called routine (last pop) is returned.
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      INTEGER FUNCTION N1RTY (IOPT,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IOPT,NFLAG
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
     
          IF (IOPT.NE.0 .AND. IOPT.NE.1) THEN
         ERTYPE(CALLVL) = 5
         ERCODE(CALLVL) = 1
         MSGLEN = 47
         CALL M1VECH ('.  The argument passed to N1RTY must be 0 or '//
     &                '1. ', MSGLEN, MSGSAV, MSGLEN)
         CALL E1PRT(NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
C         STOP
        NFLAG=84
        GOTO 1000
      ELSE
         N1RTY = ERTYPE(CALLVL+IOPT)
      END IF
C
1000  RETURN
      END
