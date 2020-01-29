C-----------------------------------------------------------------------
C  IMSL Name:  E1USR
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    November 2, 1984
C
C  Purpose:    Set USER CODE switch.
C
C  Usage:      CALL E1USR(SWITCH)
C
C  Arguments:
C     SWITCH - Character string.  (Input)
C                'ON'  Indicates that USER CODE mode is being entered.
C                'OFF' Indicates that USER CODE mode is being exited.
C  Remarks:
C     When E1POP is called from a routine while in USER CODE mode,
C     then an error message of type 1-4 will be printed (if an error
C     condition is in effect and the print table allows it).
C     However, an error message of type 1-4 will never be printed
C     if USER CODE mode is not in effect.
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1USR (SWITCH, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      CHARACTER  SWITCH*(*)
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    IFINIT
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
      EXTERNAL   E1INIT, E1MES, E1STL
C
      DATA IFINIT/0/
C                                  INITIALIZE ERROR TABLE IF NECESSARY
      IF (IFINIT .EQ. 0) THEN
         CALL E1INIT
         IFINIT = 1
      END IF
      IF (SWITCH.EQ.'ON' .OR. SWITCH.EQ.'on') THEN
         ISUSER(CALLVL) = .TRUE.
      ELSE IF (SWITCH.EQ.'OFF' .OR. SWITCH.EQ.'off') THEN
         ISUSER(CALLVL) = .FALSE.
      ELSE
         CALL E1STL (1, SWITCH)
         CALL E1MES (5, 1, 'Invalid value for SWITCH in call to'//
     &               ' E1USR.  SWITCH must be set to ''ON'' or '//
     &               '''OFF''.  SWITCH = ''%(L1)'' ', nflag)
      END IF
C
      RETURN
      END
