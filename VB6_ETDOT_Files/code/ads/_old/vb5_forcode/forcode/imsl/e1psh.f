C-----------------------------------------------------------------------
C  IMSL Name:  E1PSH
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 2, 1984
C
C  Purpose:    To push a subroutine name onto the error control stack.
C
C  Usage:      CALL E1PSH(NAME)
C
C  Arguments:
C     NAME   - A character string of length six specifing the name of
C              the subroutine.  (Input)
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1PSH (NAME,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      CHARACTER  NAME*(*)
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    IFINIT,NFLAG
      SAVE       IFINIT
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
      EXTERNAL   E1INIT, E1MES, E1STI
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KST
      INTEGER    I1KST
C
      DATA IFINIT/0/
C                                  INITIALIZE ERROR TABLE IF NECESSARY

      IF (IFINIT .EQ. 0) THEN
         CALL E1INIT
         IFINIT = 1
      END IF
      IF (CALLVL .GE. MAXLEV) THEN
         CALL E1STI (1, MAXLEV)
         CALL E1MES (5, 1, 'Error condition in E1PSH.  Push would '//
     &               'cause stack level to exceed %(I1). ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         NFLAG=89
         GOTO 1000
c         STOP
      ELSE
C                                  STORE ALLOCATION LEVEL
         IALLOC(CALLVL) = I1KST(1,NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
C                                  INCREMENT THE STACK POINTER BY ONE
         CALLVL = CALLVL + 1
C                                  PUT SUBROUTINE NAME INTO STACK
         RNAME(CALLVL) = NAME
C                                  SET ERROR TYPE AND ERROR CODE
         ERTYPE(CALLVL) = 0
         ERCODE(CALLVL) = 0
      END IF
C
1000  RETURN
      END
