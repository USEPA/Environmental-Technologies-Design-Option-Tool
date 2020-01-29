C-----------------------------------------------------------------------
C  IMSL Name:  E1STI
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 6, 1984
C
C  Purpose:    To store an integer for subsequent use within an error
C              message.
C
C  Usage:      CALL E1STI(II, IVALUE)
C
C  Arguments:
C     II     - Integer specifying the substitution index.  II must be
C              between 1 and 9.  (Input)
C     IVALUE - The integer to be stored.  (Input)
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1STI (II, IVALUE)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    II, IVALUE
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IBEG, IER, ILEN
      CHARACTER  ARRAY(14)
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    IFINIT
      CHARACTER  BLANK(1)
      SAVE       BLANK, IFINIT
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
      EXTERNAL   C1TIC, E1INIT, E1INPL
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1ERIF
      INTEGER    I1ERIF
C
      DATA BLANK/' '/, IFINIT/0/
C                                  INITIALIZE IF NECESSARY
      IF (IFINIT .EQ. 0) THEN
         CALL E1INIT
         IFINIT = 1
      END IF
      CALL C1TIC (IVALUE, ARRAY, 14, IER)
      IBEG = I1ERIF(ARRAY,14,BLANK,1)
      IF (II.GE.1 .AND. II.LE.9 .AND. IER.EQ.0) THEN
         ILEN = 15 - IBEG
         CALL E1INPL ('I', II, ILEN, ARRAY(IBEG))
      END IF
C
      RETURN
      END
