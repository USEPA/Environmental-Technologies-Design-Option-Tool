C-----------------------------------------------------------------------
C  IMSL Name:  E1PRT
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 14, 1984
C
C  Purpose:    To print an error message.
C
C  Usage:      CALL E1PRT
C
C  Arguments:  None
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1PRT(NFLAG)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    ALL, I, IBEG, IBLOC, IBLOC2, IEND, IER, IHDR, J,
     &           LERTYP, LOC, LOCM1, LOCX, MAXLOC, MAXTMP, MLOC, MOD,
     &           NCBEG, NLOC, NOUT,NFLAG
      CHARACTER  MSGTMP(70), STRING(10)
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      CHARACTER  ATLINE(9), BLANK(1), DBB(3), FROM(6), MSGTYP(8,7),
     &           PERSLA(2), QMARK, UNKNOW(8)
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
C                              SPECIFICATIONS FOR COMMON /ERCOM8/
      INTEGER    PROLVL, XXLINE(10), XXPLEN(10), ICALOC(10), INALOC(10)
      COMMON     /ERCOM8/ PROLVL, XXLINE, XXPLEN, ICALOC, INALOC
      SAVE       /ERCOM8/
C                              SPECIFICATIONS FOR COMMON /ERCOM9/
      CHARACTER  XXPROC(10)*31
      COMMON     /ERCOM9/ XXPROC
      SAVE       /ERCOM9/
      SAVE       ATLINE, BLANK, DBB, FROM, MSGTYP, PERSLA, QMARK,
     &           UNKNOW
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  MIN0
      INTRINSIC  MIN0
      INTEGER    MIN0
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   C1TIC, M1VE, UMACH
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1DX, I1ERIF
      INTEGER    I1DX, I1ERIF
C
      DATA MSGTYP/'N', 'O', 'T', 'E', ' ', ' ', ' ', ' ', 'A',
     &     'L', 'E', 'R', 'T', ' ', ' ', ' ', 'W', 'A', 'R',
     &     'N', 'I', 'N', 'G', ' ', 'F', 'A', 'T', 'A', 'L',
     &     ' ', ' ', ' ', 'T', 'E', 'R', 'M', 'I', 'N', 'A',
     &     'L', 'W', 'A', 'R', 'N', 'I', 'N', 'G', ' ', 'F',
     &     'A', 'T', 'A', 'L', ' ', ' ', ' '/
      DATA UNKNOW/'U', 'N', 'K', 'N', 'O', 'W', 'N', ' '/
      DATA ATLINE/' ', 'a', 't', ' ', 'l', 'i', 'n', 'e', ' '/
      DATA BLANK/' '/, FROM/' ', 'f', 'r', 'o', 'm', ' '/
      DATA DBB/'.', ' ', ' '/, PERSLA/'%', '/'/
      DATA QMARK/'?'/
C
      IF (MSGLEN .LE. 0) RETURN
      CALL UMACH (2, NOUT,NFLAG)
      IF (NFLAG.NE.0) GOTO 1000
      MAXTMP = 70
      MOD = 0
      LERTYP = ERTYPE(CALLVL)
      IHDR = HDRFMT(LERTYP)
      IF (IHDR .EQ. 3) THEN
         IF (XXPROC(PROLVL)(1:1).EQ.QMARK .AND. XXLINE(PROLVL).EQ.0)
     &       THEN
            IHDR = 1
         END IF
      END IF
      IEND = 0
      IF (IHDR.EQ.1 .AND. ERTYPE(CALLVL).LE.4) THEN
         MSGTMP(1) = BLANK(1)
         IEND = 1
C                                  CONVERT ERROR CODE INTO CHAR STRING
         CALL C1TIC (ERCODE(CALLVL), STRING, 10, IER)
C                                  LOCATE START OF NON-BLANK CHARACTERS
         IBEG = I1ERIF(STRING,10,BLANK,1)
C                                  M1VE IT TO MSGTMP
         CALL M1VE (STRING, IBEG, 10, 10, MSGTMP, IEND+1,
     &              IEND+11-IBEG, MAXTMP, IER)
         IEND = IEND + 11 - IBEG
      END IF
      IF (IHDR .NE. 2) THEN
         CALL M1VE (FROM, 1, 6, 6, MSGTMP, IEND+1, IEND+6, MAXTMP, IER)
         IEND = IEND + 6
      END IF
      IF (IHDR .EQ. 3) THEN
C                                  THIS IS A PROTRAN RUN TIME ERROR MSG.
C                                  RETRIEVE THE PROCEDURE NAME
         CALL M1VE (XXPROC(PROLVL), 1, XXPLEN(PROLVL), 31, MSGTMP,
     &              IEND+1, IEND+XXPLEN(PROLVL), MAXTMP, IER)
         MLOC = IEND + XXPLEN(PROLVL) + 1
         MSGTMP(MLOC) = BLANK(1)
         IEND = IEND + I1DX(MSGTMP(IEND+1),XXPLEN(PROLVL)+1,BLANK,1) -
     &          1
         IF (XXLINE(PROLVL) .GT. 0) THEN
C                                  INSERT ATLINE
            CALL M1VE (ATLINE, 1, 9, 9, MSGTMP, IEND+1, IEND+9,
     &                 MAXTMP, IER)
            IEND = IEND + 9
C                                  CONVERT PROTRAN GLOBAL LINE NUMBER
            CALL C1TIC (XXLINE(PROLVL), STRING, 10, IER)
C                                  LOCATE START OF NON-BLANK CHARACTERS
            IBEG = I1ERIF(STRING,10,BLANK,1)
C                                  M1VE GLOBAL LINE NUMBER TO MSGTMP
            CALL M1VE (STRING, IBEG, 10, 10, MSGTMP, IEND+1,
     &                 IEND+11-IBEG, MAXTMP, IER)
            IEND = IEND + 11 - IBEG
         END IF
      ELSE
C                                  THIS IS EITHER A LIBRARY ERROR MSG
C                                  OR A PROTRAN PREPROCESSOR ERROR MSG
         IF (IHDR .EQ. 1) THEN
C                                  THIS IS A LIBRARY ERROR MESSAGE.
C                                  RETRIEVE ROUTINE NAME
            CALL M1VE (RNAME(CALLVL), 1, 6, 6, MSGTMP, IEND+1, IEND+6,
     &                 MAXTMP, IER)
            MSGTMP(IEND+7) = BLANK(1)
            IEND = IEND + I1DX(MSGTMP(IEND+1),7,BLANK,1) - 1
         END IF
C                                  ADD DOT, BLANK, BLANK IF NEEDED
         IF (I1DX(MSGSAV,3,DBB,3) .NE. 1) THEN
            CALL M1VE (DBB, 1, 3, 3, MSGTMP, IEND+1, IEND+3, MAXTMP,
     &                 IER)
            IEND = IEND + 3
            MOD = 3
         END IF
      END IF
C                                  MSGTMP AND MSGSAV NOW CONTAIN THE
C                                   ERROR MESSAGE IN FINAL FORM.
      NCBEG = 59 - IEND - MOD
      ALL = 0
      IBLOC = I1DX(MSGSAV,MSGLEN,PERSLA,2)
      IF (IBLOC.NE.0 .AND. IBLOC.LT.NCBEG) THEN
         LOCM1 = IBLOC - 1
         LOC = IBLOC + 1
      ELSE IF (MSGLEN .LE. NCBEG) THEN
         LOCM1 = MSGLEN
         ALL = 1
      ELSE
         LOC = NCBEG
C                                  CHECK FOR APPROPRIATE PLACE TO SPLIT
   10    CONTINUE
         IF (MSGSAV(LOC) .NE. BLANK(1)) THEN
            LOC = LOC - 1
            IF (LOC .GT. 1) GO TO 10
            LOC = NCBEG + 1
         END IF
         LOCM1 = LOC - 1
      END IF
C                                  NO BLANKS FOUND IN FIRST NCBEG CHARS
      IF (LERTYP.GE.1 .AND. LERTYP.LE.7) THEN
c         WRITE (NOUT,99995) (MSGTYP(I,LERTYP),I=1,8),
c     &                     (MSGTMP(I),I=1,IEND), (MSGSAV(I),I=1,LOCM1)
          CONTINUE
      ELSE
c         WRITE (NOUT,99995) (UNKNOW(I),I=1,8), (MSGTMP(I),I=1,IEND),
c     &                     (MSGSAV(I),I=1,LOCM1)
        CONTINUE
      END IF
      IF (ALL .EQ. 0) THEN
C                                  PREPARE TO WRITE CONTINUATION OF
C                                    MESSAGE
C
C                                  FIND WHERE TO BREAK MESSAGE
C                                    LOC = NUMBER OF CHARACTERS OF
C                                          MESSAGE WRITTEN SO FAR
   20    LOCX = LOC + 64
         NLOC = LOC + 1
         IBLOC2 = IBLOC
         MAXLOC = MIN0(MSGLEN-LOC,64)
         IBLOC = I1DX(MSGSAV(NLOC),MAXLOC,PERSLA,2)
         IF (MSGSAV(NLOC).EQ.BLANK(1) .AND. IBLOC2.EQ.0) NLOC = NLOC +
     &       1
         IF (IBLOC .GT. 0) THEN
C                                  PAGE BREAK FOUND AT IBLOC
            LOCX = NLOC + IBLOC - 2
c            WRITE (NOUT,99996) (MSGSAV(I),I=NLOC,LOCX)
            LOC = NLOC + IBLOC
            GO TO 20
C                                  DON'T BOTHER LOOKING FOR BLANK TO
C                                    BREAK AT IF LOCX .GE. MSGLEN
         ELSE IF (LOCX .LT. MSGLEN) THEN
C                                  CHECK FOR BLANK TO BREAK THE LINE
   30       CONTINUE
            IF (MSGSAV(LOCX) .EQ. BLANK(1)) THEN
C                                  BLANK FOUND AT LOCX
c               WRITE (NOUT,99996) (MSGSAV(I),I=NLOC,LOCX)
               LOC = LOCX
               GO TO 20
            END IF
            LOCX = LOCX - 1
            IF (LOCX .GT. NLOC) GO TO 30
            LOCX = LOC + 64
C                                  NO BLANKS FOUND IN NEXT 64 CHARS
c            WRITE (NOUT,99996) (MSGSAV(I),I=NLOC,LOCX)
            LOC = LOCX
            GO TO 20
         ELSE
C                                  ALL THE REST WILL FIT ON 1 LINE
            LOCX = MSGLEN
c            WRITE (NOUT,99996) (MSGSAV(I),I=NLOC,LOCX)
         END IF
      END IF
C                                  SET LENGTH OF MSGSAV AND PLEN
C                                    TO SHOW THAT MESSAGE HAS
C                                    ALREADY BEEN PRINTED
 9000 MSGLEN = 0
      PLEN = 1
      IF (TRACON(LERTYP).EQ.1 .AND. CALLVL.GT.2) THEN
C                                  INITIATE TRACEBACK
c         WRITE (NOUT,99997)
           CONTINUE
         DO 9005  J=CALLVL, 1, -1
            IF (J .GT. 1) THEN
               IF (ISUSER(J-1)) THEN
c                  WRITE (NOUT,99998) RNAME(J), ERTYPE(J), ERCODE(J)
                        CONTINUE
               ELSE
c                  WRITE (NOUT,99999) RNAME(J), ERTYPE(J), ERCODE(J)
                  CONTINUE
               END IF
            ELSE
c               WRITE (NOUT,99998) RNAME(J), ERTYPE(J), ERCODE(J)
                CONTINUE
            END IF
 9005    CONTINUE
      END IF
C
1000  RETURN
99995 FORMAT (/, ' *** ', 8A1, ' ERROR', 59A1)
99996 FORMAT (' *** ', 9X, 64A1)
99997 FORMAT (14X, 'Here is a traceback of subprogram calls',
     &       ' in reverse order:', /, 14X, '      Routine    Error ',
     &       'type    Error code', /, 14X, '      -------    ',
     &       '----------    ----------')
99998 FORMAT (20X, A6, 5X, I6, 8X, I6)
99999 FORMAT (20X, A6, 5X, I6, 8X, I6, 4X, '(Called internally)')
      END
