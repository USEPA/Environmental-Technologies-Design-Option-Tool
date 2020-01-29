C-----------------------------------------------------------------------
C  IMSL Name:  E1MES
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 2, 1984
C
C  Purpose:    Set an error state for the current level in the stack.
C              The message is printed immediately if the error type is
C              5, 6, or 7 and the print attribute for that type is YES.
C
C  Usage:      CALL E1MES(IERTYP,IERCOD,MSGPKD)
C
C  Arguments:
C     IERTYP - Integer specifying the error type.  (Input)
C                IERTYP=1,  informational/note
C                IERTYP=2,  informational/alert
C                IERTYP=3,  informational/warning
C                IERTYP=4,  informational/fatal
C                IERTYP=5,  terminal
C                IERTYP=6,  PROTRAN/warning
C                IERTYP=7,  PROTRAN/fatal
C     IERCOD - Integer specifying the error code.  (Input)
C     MSGPKD - A character string containing the message.
C              (Input)  Within the message, any of following may appear
C                %(A1),%(A2),...,%(A9) for character arrays
C                %(C1),%(C2),...,%(C9) for complex numbers
C                %(D1),%(D2),...,%(D9) for double precision numbers
C                %(I1),%(I2),...,%(I9) for integer numbers
C                %(K1),%(K2),...,%(K9) for keywords
C                %(L1),%(L2),...,%(L9) for literals (strings)
C                %(R1),%(R2),...,%(R9) for real numbers
C                %(Z1),%(Z2),...,%(Z9) for double complex numbers
C              This provides a way to insert character arrays, strings,
C              numbers, and keywords into the message.  See remarks
C              below.
C
C  Remarks:
C     The number of characters in the message after the insertion of
C     the corresponding strings, etc. should not exceed 255.  If the
C     limit is exceeded, only the first 255 characters will be used.
C     The appropriate strings, etc. need to have been previously stored
C     in common via calls to E1STA, E1STD, etc.  Line breaks may be
C     specified by inserting the two characters '%/' into the message
C     at the desired locations.
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1MES (IERTYP, IERCOD, MSGPKD,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IERTYP, IERCOD, NFLAG
      CHARACTER  MSGPKD*(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    ERTYP2, I, IER, IPLEN, ISUB, LAST, LEN2, LOC, M, MS,
     &           NLOC, NUM, PBEG
      CHARACTER  MSGTMP(255)
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    IFINIT, NFORMS
      CHARACTER  BLNK, DBB(3), FIND(4), FORMS(9), INREF(25), LPAR,
     &           NCHECK(3), PERCNT, RPAR
      SAVE       BLNK, DBB, FIND, FORMS, IFINIT, INREF, LPAR, NCHECK,
     &           NFORMS, PERCNT, RPAR
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
C     INTRINSIC  LEN,MIN0
      INTRINSIC  LEN, MIN0
      INTEGER    LEN, MIN0
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   C1TCI, E1INIT, E1PRT, E1UCS, M1VE, M1VECH
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1DX
      INTEGER    I1DX
C
      DATA FORMS/'A', 'C', 'D', 'I', 'K', 'L', 'R', 'S', 'Z'/,
     &     NFORMS/9/
      DATA PERCNT/'%'/, LPAR/'('/, RPAR/')'/, BLNK/' '/
      DATA INREF/' ', 'i', 'n', ' ', 'r', 'e', 'f', 'e', 'r',
     &     'e', 'n', 'c', 'e', ' ', 't', 'o', ' ', 'k', 'e',
     &     'y', 'w', 'o', 'r', 'd', ' '/
      DATA NCHECK/'N', '1', '*'/, DBB/'.', ' ', ' '/
      DATA FIND/'*', ' ', ' ', '*'/
      DATA IFINIT/0/
    
C                                  INITIALIZE ERROR TABLE IF NECESSARY
      IF (IFINIT .EQ. 0) THEN
         CALL E1INIT
         IFINIT = 1
      END IF
C                                  CHECK AND SET ERROR TYPE IF NECESSARY
      IF (IERTYP .NE. -1) THEN
         ERTYPE(CALLVL) = IERTYP
      ELSE IF (IERTYP.LT.-1 .OR. IERTYP.GT.7) THEN
         MSGLEN = 51
         CALL M1VECH ('.  Error from E1MES.  Illegal error type'//
     &                ' specified. ', MSGLEN, MSGSAV, MSGLEN)
         CALL E1PRT(NFLAG)
         NFLAG=98
         IF (NFLAG.NE.0) GOTO 1000
c         STOP
      END IF
C
      ERTYP2 = ERTYPE(CALLVL)
C                                  SET ERROR CODE IF NECESSARY
      IF (IERCOD .GT. -1) ERCODE(CALLVL) = IERCOD
      LEN2 = LEN(MSGPKD)
C
      IF (IERTYP.EQ.0 .OR. IERCOD.EQ.0) THEN
C                                  REMOVE THE ERROR STATE
         MSGLEN = 0
      ELSE IF (LEN2.EQ.0 .OR. (LEN2.EQ.1.AND.MSGPKD(1:1).EQ.BLNK)) THEN
         IF (ERTYP2 .EQ. 6) IFERR6 = 1
         IF (ERTYP2 .EQ. 7) IFERR7 = 1
C                                  UPDATE CHECKSUM PARAMETER ERCKSM
         CALL E1UCS
C                                  PRINT MESSAGE IF NECESSARY
         IF (ERTYP2.GE.5 .AND. PRINTB(ERTYP2).EQ.1) THEN
                      CALL E1PRT(NFLAG) 
                      IF (NFLAG.NE.0) GOTO 1000
                      ENDIF
      ELSE
C                                  FILL UP MSGSAV WITH EXPANDED MESSAGE
         LEN2 = MIN0(LEN2,255)
         DO 10  I=1, LEN2
            MSGTMP(I) = MSGPKD(I:I)
   10    CONTINUE
         MS = 0
         M = 0
C                                  CHECK PLIST FOR KEYWORD NAME
         NLOC = I1DX(PLIST,PLEN,NCHECK,3)
         IF (NLOC.GT.0 .AND. HDRFMT(ERTYP2).EQ.3) THEN
C                                  M1VE INREF INTO MSGSAV
            CALL M1VE (INREF, 1, 25, 25, MSGSAV, 1, 25, 25, IER)
C                                  GET LENGTH OF KEYWORD NAME
            CALL C1TCI (PLIST(NLOC+3), 3, IPLEN, IER,NFLAG)
            IF (NFLAG.NE.0) GOTO 1000
            PBEG = NLOC + 3 + IER
C                                  M1VE KEYWORD NAME INTO MSGSAV
            CALL M1VE (PLIST, PBEG, PBEG+IPLEN-1, PLEN, MSGSAV, 26,
     &                 IPLEN+25, 255, IER)
C                                  UPDATE POINTER
            MS = IPLEN + 25
         END IF
C                                  INSERT DOT, BLANK, BLANK
         CALL M1VE (DBB, 1, 3, 3, MSGSAV, MS+1, MS+3, 255, IER)
         MS = MS + 3
C                                  LOOK AT NEXT CHARACTER
   20    M = M + 1
         ISUB = 0
         IF (M .GT. LEN2-4) THEN
            LAST = LEN2 - M + 1
            DO 30  I=1, LAST
   30       MSGSAV(MS+I) = MSGTMP(M+I-1)
            MSGLEN = MS + LAST
            GO TO 40
         ELSE IF (MSGTMP(M).EQ.PERCNT .AND. MSGTMP(M+1).EQ.LPAR .AND.
     &           MSGTMP(M+4).EQ.RPAR) THEN
            CALL C1TCI (MSGTMP(M+3), 1, NUM, IER,NFLAG)
            IF (NFLAG.NE.0) GOTO 1000
            IF (IER.EQ.0 .AND. NUM.NE.0 .AND. I1DX(FORMS,NFORMS,
     &          MSGTMP(M+2),1).NE.0) THEN
C                                  LOCATE THE ITEM IN THE PARAMETER LIST
               CALL M1VE (MSGTMP(M+2), 1, 2, 2, FIND, 2, 3, 4, IER)
               LOC = I1DX(PLIST,PLEN,FIND,4)
               IF (LOC .GT. 0) THEN
C                                  SET IPLEN = LENGTH OF STRING
                  CALL C1TCI (PLIST(LOC+4), 4, IPLEN, IER,NFLAG)
                  IF (NFLAG.NE.0) GOTO 1000
                  PBEG = LOC + 4 + IER
C                                  ADJUST IPLEN IF IT IS TOO BIG
                  IPLEN = MIN0(IPLEN,255-MS)
C                                  M1VE STRING FROM PLIST INTO MSGSAV
                  CALL M1VE (PLIST, PBEG, PBEG+IPLEN-1, PLEN, MSGSAV,
     &                       MS+1, MS+IPLEN, 255, IER)
                  IF (IER.GE.0 .AND. IER.LT.IPLEN) THEN
C                                  UPDATE POINTERS
                     M = M + 4
                     MS = MS + IPLEN - IER
C                                  BAIL OUT IF NO MORE ROOM
                     IF (MS .GE. 255) THEN
                        MSGLEN = 255
                        GO TO 40
                     END IF
C                                  SET FLAG TO SHOW SUBSTITION WAS MADE
                     ISUB = 1
                  END IF
               END IF
            END IF
         END IF
         IF (ISUB .EQ. 0) THEN
            MS = MS + 1
            MSGSAV(MS) = MSGTMP(M)
         END IF
         GO TO 20
   40    ERTYP2 = ERTYPE(CALLVL)
         IF (ERTYP2 .EQ. 6) IFERR6 = 1
         IF (ERTYP2 .EQ. 7) IFERR7 = 1
C                                  UPDATE CHECKSUM PARAMETER ERCKSM
         CALL E1UCS
C                                  PRINT MESSAGE IF NECESSARY
         IF (ERTYP2.GE.5 .AND. PRINTB(ERTYP2).EQ.1) THEN 
               CALL E1PRT(NFLAG)
               IF (NFLAG.NE.0) GOTO 1000
               ENDIF
      END IF
C                                  CLEAR PARAMETER LIST
      PLEN = 1
C
1000  RETURN
      END
