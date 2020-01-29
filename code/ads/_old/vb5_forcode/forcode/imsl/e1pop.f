C-----------------------------------------------------------------------
C  IMSL Name:  E1POP
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 13, 1984
C
C  Purpose:    To pop a subroutine name from the error control stack.
C
C  Usage:      CALL E1POP(NAME)
C
C  Arguments:
C     NAME   - A character string of length six specifying the name
C              of the subroutine.  (Input)
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE E1POP (NAME,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      CHARACTER  NAME*(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IERTYP, IR,NFLAG,NFF
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
      EXTERNAL   E1MES, E1PRT, E1PSH, E1STI, E1STL, I1KRL
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KST
      INTEGER    I1KST
C

      IF (CALLVL .LE. 1) THEN
         CALL E1PSH ('E1POP ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1STL (1, NAME)
         CALL E1MES (5, 1, 'Error condition in E1POP.  Cannot pop '//
     &               'from %(L1) because stack is empty.',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         NFLAG=88
         GOTO 1000
c         STOP
      ELSE IF (NAME .NE. RNAME(CALLVL)) THEN
         CALL E1STL (1, NAME)
         CALL E1STL (2, RNAME(CALLVL))
         CALL E1MES (5, 2, 'Error condition in E1POP.  %(L1) does '//
     &           'not match the name %(L2) in the stack.',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         NFLAG=87
         GOTO 1000
         
c         STOP
      ELSE
         IERTYP = ERTYPE(CALLVL)
         IF (IERTYP .NE. 0) THEN
C                                  M1VE ERROR TYPE AND ERROR CODE TO
C                                    PREVIOUS LEVEL FOR ERROR TYPES 2-7
            IF (IERTYP.GE.2 .AND. IERTYP.LE.7) THEN
               ERTYPE(CALLVL-1) = ERTYPE(CALLVL)
               ERCODE(CALLVL-1) = ERCODE(CALLVL)
            END IF
C                                  CHECK PRINT TABLE TO DETERMINE
C                                    WHETHER TO PRINT STORED MESSAGE
            IF (IERTYP .LE. 4) THEN
               IF (ISUSER(CALLVL-1) .AND. PRINTB(IERTYP).EQ.1)
     &            THEN 
                        CALL E1PRT(NFLAG)
                        IF (NFLAG.NE.0) GOTO 1000
                        ENDIF
            ELSE
               IF (PRINTB(IERTYP) .EQ. 1) THEN 
                  CALL E1PRT(NFLAG)
                  IF (NFLAG.NE.0) GOTO 1000
                  ENDIF

            END IF
C                                  CHECK STOP TABLE AND ERROR TYPE TO
C                                    DETERMINE WHETHER TO STOP
            IF (IERTYP .LE. 4) THEN
               IF (ISUSER(CALLVL-1) .AND. STOPTB(IERTYP).EQ.1) THEN
                  NFLAG=70
                  GOTO 1000
               END IF
            ELSE IF (IERTYP .EQ. 5) THEN
               IF (STOPTB(IERTYP) .EQ. 1) THEN
                  NFLAG=71
                  GOTO 1000
               END IF
            ELSE IF (HDRFMT(IERTYP) .EQ. 1) THEN
               IF (ISUSER(CALLVL-1)) THEN
                  IF (N1RGB(0) .NE. 0) THEN
                  NFLAG=72
                  GOTO 1000
                  END IF
               END IF
            END IF
         END IF
C                                  SET ERROR TYPE AND CODE
         IF (CALLVL .LT. MAXLEV) THEN
            ERTYPE(CALLVL+1) = -1
            ERCODE(CALLVL+1) = -1
         END IF
C                                  SET IR = AMOUNT OF WORKSPACE
C                                  ALLOCATED AT THIS LEVEL
         IR = I1KST(1,NFLAG) - IALLOC(CALLVL-1)
         IF (NFLAG.NE.0) GOTO 1000
         IF (IR .GT. 0) THEN
C                                  RELEASE WORKSPACE
            CALL I1KRL (IR,NFLAG)
            IF (NFLAG.NE.0) GOTO 1000
            IALLOC(CALLVL) = 0
         ELSE IF (IR .LT. 0) THEN
            CALL E1STI (1, CALLVL)
            CALL E1STI (2, IALLOC(CALLVL-1))
            NFF=I1KST(1,NFLAG)
            IF (NFLAG.NE.0) GOTO 1000
            CALL E1STI (3, NFF)
            CALL E1MES (5, 3, 'Error condition in E1POP. '//
     &                  ' The number of workspace allocations at '//
     &                  'level %(I1) is %(I2).  However, the total '//
     &         'number of workspace allocations is %(I3).',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         NFLAG=86
         GOTO 1000
            
c            STOP
         END IF
C                                  DECREASE THE STACK POINTER BY ONE
         CALLVL = CALLVL - 1
      END IF
C
1000  RETURN
      END
