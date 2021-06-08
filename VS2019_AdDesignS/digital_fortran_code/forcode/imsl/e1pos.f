C-----------------------------------------------------------------------
C  IMSL Name:  E1POS
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 2, 1984
C
C  Purpose:    Set or retrieve print and stop attributes.
C
C  Usage:      CALL E1POS(IERTYP,IPATT,ISATT)
C
C  Arguments:
C     IERTYP - Integer specifying the error type for which print and
C              stop attributes are to be set or retrieved.  (Input)  If
C              IERTYP is 0 then the settings apply to all error types.
C              If IERTYP is between 1 and 7, then the settings only
C              apply to that specified error type.  If IERTYP is
C              negative then the current print and stop attributes will
C              be returned in IPATT and ISATT.
C     IPATT  - If IERTYP is positive, IPATT is an integer specifying the
C              desired print attribute as follows: -1 means no change,
C              0 means NO, 1 means YES, and 2 means assign the default
C              setting.  (Input)  If IERTYP is negative, IPATT is
C              returned as 1 if print is YES or 0 if print is NO for
C              error type IABS(IERTYP).  (Output)
C     ISATT  - If IERTYP is positive, ISATT is an integer specifying the
C              desired stop attribute as follows: -1 means no change,
C              0 means NO, 1 means YES, and 2 means assign the default
C              setting.  (Input)  If IERTYP is negative, ISATT is
C              returned as 1 if print is YES or 0 if print is NO for
C              error type IABS(IERTYP).  (Output)
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1POS (IERTYP, IPATT, ISATT,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IERTYP, IPATT, ISATT,NFLAG
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IER
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    DEFLTP(7), DEFLTS(7), IFINIT
      SAVE       DEFLTP, DEFLTS, IFINIT
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
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  IABS
      INTRINSIC  IABS
      INTEGER    IABS
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1INIT, E1MES, E1STI
C
      DATA IFINIT/0/
      DATA DEFLTP/0, 0, 1, 1, 1, 1, 1/, DEFLTS/0, 0, 0, 1, 1, 0, 1/
C                                  INITIALIZE ERROR TABLE IF NECESSARY
        
      IF (IFINIT .EQ. 0) THEN
         CALL E1INIT
         IFINIT = 1
      END IF
      IER = 0
      IF (IERTYP .GE. 0) THEN
         IF (IPATT.LT.-1 .OR. IPATT.GT.2) THEN
            CALL E1STI (1, IPATT)
            CALL E1MES (5, 1, 'Invalid value specified for print '//
     &                  'table attribute.  IPATT must be -1, 0, 1, '//
     &                  'or 2.  IPATT = %(I1)',NFLAG)
            IF (NFLAG.NE.0) GOTO 1000
            IER = 1
         END IF
         IF (ISATT.LT.-1 .OR. ISATT.GT.2) THEN
            CALL E1STI (1, ISATT)
            CALL E1MES (5, 1, 'Invalid value specified for stop '//
     &                  'table attribute.  ISATT must be -1, 0, 1, '//
     &                  'or 2.  ISATT = %(I1)',NFLAG)
            IF (NFLAG.NE.0) GOTO 1000
            IER = 1
         END IF
      END IF
      IF (IER .EQ. 0) THEN
         IF (IERTYP .EQ. 0) THEN
            IF (IPATT.EQ.0 .OR. IPATT.EQ.1) THEN
               DO 10  I=1, 7
   10          PRINTB(I) = IPATT
            ELSE IF (IPATT .EQ. 2) THEN
C                                  ASSIGN DEFAULT SETTINGS
               DO 20  I=1, 7
   20          PRINTB(I) = DEFLTP(I)
            END IF
            IF (ISATT.EQ.0 .OR. ISATT.EQ.1) THEN
               DO 30  I=1, 7
   30          STOPTB(I) = ISATT
            ELSE IF (ISATT .EQ. 2) THEN
C                                  ASSIGN DEFAULT SETTINGS
               DO 40  I=1, 7
   40          STOPTB(I) = DEFLTS(I)
            END IF
         ELSE IF (IERTYP.GE.1 .AND. IERTYP.LE.7) THEN
            IF (IPATT.EQ.0 .OR. IPATT.EQ.1) THEN
               PRINTB(IERTYP) = IPATT
            ELSE IF (IPATT .EQ. 2) THEN
C                                  ASSIGN DEFAULT SETTING
               PRINTB(IERTYP) = DEFLTP(IERTYP)
            END IF
            IF (ISATT.EQ.0 .OR. ISATT.EQ.1) THEN
               STOPTB(IERTYP) = ISATT
            ELSE IF (ISATT .EQ. 2) THEN
C                                  ASSIGN DEFAULT SETTING
               STOPTB(IERTYP) = DEFLTS(IERTYP)
            END IF
         ELSE IF (IERTYP.LE.-1 .AND. IERTYP.GE.-7) THEN
            I = IABS(IERTYP)
            IPATT = PRINTB(I)
            ISATT = STOPTB(I)
         END IF
      END IF
C
1000  RETURN
      END
