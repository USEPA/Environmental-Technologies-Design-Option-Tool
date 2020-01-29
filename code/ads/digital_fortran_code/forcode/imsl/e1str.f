C-----------------------------------------------------------------------
C  IMSL Name:  E1STR
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 2, 1984
C
C  Purpose:    To store a real number for subsequent use within an error
C              message.
C
C  Usage:      CALL E1STR(IR,RVALUE)
C
C  Arguments:
C     IR     - Integer specifying the substitution index.  IR must be
C              between 1 and 9.  (Input)
C     RVALUE - The real number to be stored.  (Input)
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1STR (IR, RVALUE)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IR
      REAL       RVALUE
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IBEG, ILEN
      CHARACTER  ARRAY(14), SAVE*14
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
      EXTERNAL   E1INIT, E1INPL
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
      IF (RVALUE .EQ. 0.0) THEN
         WRITE (SAVE,'(E14.6)') RVALUE
      ELSE
         WRITE (SAVE,'(1PE14.6)') RVALUE
      END IF
      DO 40  I=1, 14
   40 ARRAY(I) = SAVE(I:I)
      IBEG = I1ERIF(ARRAY,14,BLANK,1)
      IF (IR.GE.1 .AND. IR.LE.9) THEN
         ILEN = 15 - IBEG
         CALL E1INPL ('R', IR, ILEN, ARRAY(IBEG))
      END IF
C
      RETURN
      END
