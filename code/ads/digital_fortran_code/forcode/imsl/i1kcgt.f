C-----------------------------------------------------------------------
C  IMSL Name:  I1KCGT
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    January 31, 1984
C
C  Purpose:    Allocate character workspace out of the array CHWKSP
C              in COMMON block WKSPCH for a CHARACTER*(LEN) array of
C              length NELMTS.
C
C  Usage:      I1KCGT (NELMTS, LEN)
C
C  Arguments:
C     NELMTS - Number of elements of a specified character length
C              to be allocated.  (Input)
C     LEN    - Desired length of the individual character allocations.
C              (Input)
C     I1KCGT - Integer function.  (Output)  Returns the index of the
C              first element in the current allocation.
C
C  Remarks:
C  1. On return, the array will occupy C?WSP(I1KCGT), C?WSP(I1KCGT+1),
C     ..., C?WSP(I1KCGT+NELMTS-1) where C?WSP is a CHARACTER*(?) array
C     equivalenced to CHWKSP.
C
C  2. If I1KCGT is negative, the absolute value of I1KCGT is the
C     additional workspace needed for the current allocation.
C
C  3. The allocator reserves the first nine integer locations of
C     permanent storage area of the stack in COMMON block RWKSP for
C     bookkeeping purposes.  These pointers are allocated for and
C     initialized by the function IWKCIN upon the first call to the
C     character allocation package.
C
C  4. By default, a maximum of 20 outstanding (unreleased) allocations
C     are allowed from a maximum character space of 2000 characters.
C     Both these limits may be increased or decreased by calling
C     IWKCIN.
C
C  5. The character bookkeeping array stored in IWKSP of COMMON block
C     RWKSP starts at index NDXS = IWKSP(9) and it is IWKSP(10) integer
C     words long.  The use of the first nine integer locations is as
C     follows:
C
C     IWKSP(NDXS  ) - LOUT    The number of current character
C                             allocations.
C     IWKSP(NDXS+1) - LNOW    The current active length of the character
C                             stack.
C     IWKSP(NDXS+2) - LUSED   The maximum value of IWKSP(NDXS+1)
C                             achieved thus far.
C     IWKSP(NDXS+3) - LBND    The lower bound of permanent character
C                             storage which is one character more than
C                             the maximum allowed length of the stack.
C     IWKSP(NDXS+4) - LMAX    The maximum length of the character
C                             storage array.
C     IWKSP(NDXS+5) - LALC    The total number of character allocations
C                             handled by I1KCGT
C     IWKSP(NDXS+6) - LNEED   The number of CHARACTER*1 units by which
C                             the array size must be increased for all
C                             past allocations to succeed.
C     IWKSP(NDXS+7) - LBOOK   The number of integer locations in the
C                             permanent portion of IWKSP which are used
C                             for bookkeeping.
C     IWKSP(NDXS+8) - LBPERM  The lower bound of the pointers kept in
C                             the permanent portion of IWKSP for the
C                             permanent character storage area of
C                             CHWKSP.
C     IWKSP(NDXS+9) - LMAXA   The maximum number of allocations allowed
C                             under the current call to IWKCIN. The
C                             default of 20 is replaced by the value of
C                             the argument NALC in subroutine IWKCIN.
C     IWKSP(NDXS+10) - LNEEDA The number by which the maximum number of
C                             allocations (LMAXA) must be increased for
C                             all past allocations to succeed.
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C

      INTEGER FUNCTION I1KCGT (NELMTS, LEN,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NELMTS, LEN,NFLAG,NFLAGBIS,II
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IDUMAL, IGAP, ILEFT, IPA, IPA7, ISA, ISA7, LALC,
     &           LBND, LBOOK, LBPERM, LEN2, LMAX, LMAXA, LNEED,
     &           LNEED1, LNEEDA, LNOW, LOUT, LUSED, NDX, NDXS
      CHARACTER  TY1*64, TY2*58, TY3*46
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    INIT
      CHARACTER  DLMTR*1
      SAVE       DLMTR, INIT
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                              SPECIFICATIONS FOR COMMON /ERCOM8/
      INTEGER    PROLVL, XXLINE(10), XXPLEN(10), ICALOC(10), INALOC(10)
      COMMON     /ERCOM8/ PROLVL, XXLINE, XXPLEN, ICALOC, INALOC
      SAVE       /ERCOM8/
C                              SPECIFICATIONS FOR COMMON /ERCOM9/
      CHARACTER  XXPROC(10)*31
      COMMON     /ERCOM9/ XXPROC
      SAVE       /ERCOM9/
C                                  SPECIFICATIONS FOR COMMON /WORKSP/
      REAL       RWKSP(5000)
      REAL       RDWKSP(5000)
      DOUBLE PRECISION DWKSP(2500)
      COMPLEX    CWKSP(2500)
      COMPLEX    CZWKSP(2500)
      COMPLEX    *16 ZWKSP(1250)
      INTEGER    IWKSP(5000)
      LOGICAL    LWKSP(5000)
      EQUIVALENCE (DWKSP(1), RWKSP(1))
      EQUIVALENCE (CWKSP(1), RWKSP(1)), (ZWKSP(1), RWKSP(1))
      EQUIVALENCE (IWKSP(1), RWKSP(1)), (LWKSP(1), RWKSP(1))
      EQUIVALENCE (RDWKSP(1), RWKSP(1)), (CZWKSP(1), RWKSP(1))
      COMMON     /WORKSP/ RWKSP
C                                  SPECIFICATIONS FOR COMMON /WKSPCH/
      CHARACTER  *1 CHWKSP(2000)
      COMMON     /WKSPCH/ CHWKSP
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  IABS,MAX0,MOD
      INTRINSIC  IABS, MAX0, MOD
      INTEGER    IABS, MAX0, MOD
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1POS, E1PSH, E1STI, I1KC00
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KCQU
      INTEGER    I1KCQU
C
      DATA INIT/.TRUE./
      DATA DLMTR/'$'/
C
      NFLAGBIS=0
      CALL E1PSH ('I1KCGT',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C                                  INITIALIZE FUNCTION VALUE TO 0
      I1KCGT = 0
C
      IF (INIT) CALL I1KC00 (INIT, 2000, 20,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      NDXS = IWKSP(9)
C                                  INITIALIZE LOCAL VARIABLES
      LOUT = IWKSP(NDXS)
      LNOW = IWKSP(NDXS+1)
      LUSED = IWKSP(NDXS+2)
      LBND = IWKSP(NDXS+3)
      LMAX = IWKSP(NDXS+4)
      LALC = IWKSP(NDXS+5)
      LNEED = IWKSP(NDXS+6)
      LBOOK = IWKSP(NDXS+7)
      LBPERM = IWKSP(NDXS+8)
      LMAXA = IWKSP(NDXS+9)
      LNEEDA = IWKSP(NDXS+10)
C                                  NUMBER OF ITEMS REQUESTED LESS THAN 0
      IF (NELMTS .LT. 0) THEN
         CALL E1STI (1, NELMTS)
         CALL E1MES (5, 1, 'The number of items is not positive. '//
     &               'NELMTS = %(I1).',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POP ('I1KCGT',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=160
         GO TO 9000
      END IF
C                                  CHARACTER LENGTH IS 0
      IF (LEN .EQ. 0) THEN
         CALL E1MES (5, 2, 'The requested length of CHARACTER '//
     &               'variable is 0.',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POP ('I1KCGT',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=161
         GO TO 9000
      END IF
C                                  CHECK POINTERS IN NUMERICAL
C                                  BOOKKEEPING SECTION
      IF (NDXS.LT.IWKSP(4) .OR. NDXS.GE.IWKSP(5)) THEN
         CALL E1MES (5, 3, 'Stack pointers have been overwritten '//
     &               'in integer array IWKSP. ',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POP ('I1KCGT',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=162
         GO TO 9000
      END IF
C                                  CHECK CHARACTER WORKSPACE POINTERS
C                                  WHICH RESIDE IN A PART OF THE
C                                  PERMANENT SECTION OF THE NUMERICAL
C                                  WORKSPACE ARRAY
      IF (LNOW.LT.0 .OR. LNOW.GT.LUSED .OR. LUSED.GT.LMAX .OR.
     &    LNOW.GE.LBND .OR. LOUT.GT.LALC) THEN
         CALL E1MES (5, 4, 'One or more of the character '//
     &               'workspace bookkeeping locations have been '//
     &               'overwritten.',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POP ('I1KCGT',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=163
         GO TO 9000
      END IF
C
      CALL E1POP ('I1KCGT',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C                                  DETERMINE NUMBER OF LOCATIONS STILL
C                                  AVAILABLE FOR CHARACTER*(LEN)
      II=I1KCQU(IABS(LEN),NFLAG)
      IF(NFLAG.NE.0) GOTO 9999
      ILEFT =II 
C      I1KCQU(IABS(LEN))
C                                  Set temporary variables to hold part
C                                  of error messages.  This is done to
C                                  so that it will compile on certain
C                                  machines.
      TY1 = 'Insufficient character workspace for the current '//
     &      'allocation(s). '
      TY2 = 'The number of current character allocations is too '//
     &      'large. '
      TY3 = 'Number of character allocations too large AND '
C
      IF (LEN .GT. 0) THEN
C                                  RELEASABLE STORAGE
         IF (ILEFT .GE. NELMTS) THEN
            IF (LOUT+1 .GT. LMAXA) THEN
               LNEEDA = LNEEDA + 1
               I1KCGT = 0
            ELSE
               I1KCGT = (LNOW+LEN-1)/LEN + 1
               I = (I1KCGT-1+NELMTS)*LEN + 1
C                                  CHWKSP(I) CONTAINS A CHARACTER TO
C                                  BE CHECKED FOR UPON DEALLOCATION
C                                  TO DETECT IF THE CHARACTER ARRAY
C                                  HAS BEEN OVERWRITTEN
               CHWKSP(I) = DLMTR
               NDX = NDXS + LBOOK + 2*LOUT
C                                  IWKSP(NDX) CONTAINS THE LENGTH OF THE
C                                  CHARACTER DATA FOR THIS ALLOCATION.
C                                  IWKSP(NDX+1) CONTAINS LNOW FOR THE
C                                  PREVIOUS ALLOCATION.
               IWKSP(NDX) = LEN
               IWKSP(NDX+1) = LNOW
C
               LOUT = LOUT + 1
               LALC = LALC + 1
               LNOW = I
               LUSED = MAX0(LUSED,I)
               LNEED = 0
            END IF
         ELSE
            IF (LOUT+1 .GT. LMAXA) LNEEDA = LNEEDA + 1
C
C                                  RELEASABLE STORAGE WAS REQUESTED BUT
C                                  THE STACK WOULD OVERFLOW. THEREFORE,
C                                  ALLOCATE RELEASABLE STORAGE THROUGH
C                                  THE END OF THE STACK.
            IF (LNEED .EQ. 0) THEN
               IDUMAL = (LNOW-1)/LEN + 2
               I = (IDUMAL-1+ILEFT)*LEN + 1
               IF (I .LT. LBND) THEN
C                                  CHWKSP(I) CONTAINS A CHARACTER TO
C                                  BE CHECKED FOR UPON DEALLOCATION
C                                  TO DETECT IF THE CHARACTER ARRAY
C                                  HAS BEEN OVERWRITTEN
                  CHWKSP(I) = DLMTR
                  NDX = NDXS + LBOOK + 2*LOUT
C                                  IWKSP(NDX) CONTAINS THE LENGTH OF THE
C                                  CHARACTER DATA FOR THIS ALLOCATION.
C                                  IWKSP(NDX+1) CONTAINS LNOW FOR THE
C                                  PREVIOUS ALLOCATION.
                  IWKSP(NDX) = LEN
                  IWKSP(NDX+1) = LNOW
C
                  LOUT = LOUT + 1
                  LALC = LALC + 1
                  LNOW = I
                  LUSED = MAX0(LUSED,I)
               END IF
            END IF
C                                  CALCULATE THE AMOUNT OF SPACE NEEDED
C                                  TO ACCOMODATE THIS ALLOCATION REQUEST
            LNEED1 = (NELMTS-ILEFT)*LEN
            IF (ILEFT .EQ. 0) THEN
               IGAP = LEN - MOD(LNOW+LNEED,LEN)
               IF (IGAP .EQ. LEN) IGAP = 0
               LNEED1 = LNEED1 + 1 + IGAP
            END IF
            LNEED = LNEED + LNEED1
            I1KCGT = -LNEED
         END IF
      ELSE
C                                  PERMANENT STORAGE
         LEN2 = -LEN
         IF (ILEFT .GE. NELMTS) THEN
            IF (LOUT+1 .GT. LMAXA) THEN
               LNEEDA = LNEEDA + 1
               I1KCGT = 0
            ELSE
               I1KCGT = (LBND-1)/LEN2 + 1 - NELMTS
               I = (I1KCGT-1)*LEN2 - 1
C                                  CHWKSP(I) CONTAINS A CHARACTER TO
C                                  BE CHECKED FOR UPON DEALLOCATION
C                                  TO DETECT IF THE CHARACTER ARRAY
C                                  HAS BEEN OVERWRITTEN
               CHWKSP(I) = DLMTR
               NDX = NDXS + LBPERM - 2
C                                  IWKSP(NDX) CONTAINS THE LENGTH OF THE
C                                  CHARACTER DATA FOR THIS ALLOCATION.
C                                  IWKSP(NDX+1) CONTAINS LNOW FOR THE
C                                  PREVIOUS ALLOCATION.
               IWKSP(NDX) = LBND
               IWKSP(NDX+1) = LEN
C
               LBND = I
               LNEED = 0
               LBPERM = LBPERM - 2
            END IF
         ELSE
            IF (LOUT+1 .GT. LMAXA) LNEEDA = LNEEDA + 1
C
C                                  PERMANENT STORAGE WAS REQUESTED BUT
C                                  THE STACK WOULD OVERFLOW. THEREFORE,
C                                  ALLOCATE RELEASABLE STORAGE THROUGH
C                                  THE END OF THE STACK.
            IF (LNEED .EQ. 0) THEN
               IDUMAL = (LNOW-1)/LEN2 + 2
               I = (IDUMAL-1+ILEFT)*LEN2 + 1
               IF (I .LT. LBND) THEN
C                                  CHWKSP(I) CONTAINS A CHARACTER TO
C                                  BE CHECKED FOR UPON DEALLOCATION
C                                  TO DETECT IF THE CHARACTER ARRAY
C                                  HAS BEEN OVERWRITTEN
                  CHWKSP(I) = DLMTR
                  NDX = NDXS + LBPERM - 2
C                                  IWKSP(NDX) CONTAINS THE LENGTH OF THE
C                                  CHARACTER DATA FOR THIS ALLOCATION.
C                                  IWKSP(NDX+1) CONTAINS LNOW FOR THE
C                                  PREVIOUS ALLOCATION.
                  IWKSP(NDX) = LEN
                  IWKSP(NDX+1) = LNOW
C
                  LOUT = LOUT + 1
                  LNOW = I
                  LALC = LALC + 1
                  LUSED = MAX0(LUSED,I)
               END IF
            END IF
C                                  CALCULATE THE AMOUNT NEEDED TO
C                                  ACCOMODATE THIS ALLOCATION REQUEST
            LNEED1 = (NELMTS-ILEFT)*LEN2
            IF (ILEFT .EQ. 0) THEN
               IGAP = LEN2 - MOD(LNOW+LNEED,LEN2)
               IF (IGAP .EQ. LEN2) IGAP = 0
               LNEED1 = LNEED1 + 1 + IGAP
            END IF
            LNEED = LNEED + LNEED1
            I1KCGT = -LNEED
         END IF
      END IF
C                                  REASSIGN VECTOR LOCATIONS TO NEWLY
C                                  CALCULATED VALUES
      IWKSP(NDXS) = LOUT
      IWKSP(NDXS+1) = LNOW
      IWKSP(NDXS+2) = LUSED
      IWKSP(NDXS+3) = LBND
      IWKSP(NDXS+4) = LMAX
      IWKSP(NDXS+5) = LALC
      IWKSP(NDXS+6) = LNEED
      IWKSP(NDXS+7) = LBOOK
      IWKSP(NDXS+8) = LBPERM
      IWKSP(NDXS+9) = LMAXA
      IWKSP(NDXS+10) = LNEEDA
C                                  STACK OVERFLOW - UNRECOVERABLE ERROR
 9000 IF (LNEED.GT.0 .OR. LNEEDA.GT.0) THEN
         CALL E1POS (-5, IPA, ISA,NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POS (5, 0, 0,NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POS (-7, IPA7, ISA7,NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POS (7, 0, 0,NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1PSH ('I1KCGT',NFLAG)
         IF(NFLAG.NE.0) GOTO 9999
         CALL E1STI (1, LNEED+LMAX)
         CALL E1STI (2, LNEEDA+LMAXA)
         IF (LNEED.GT.0 .AND. LNEEDA.EQ.0) THEN
            IF (XXLINE(PROLVL).GE.1 .AND. XXLINE(PROLVL).LE.999) THEN
               CALL E1MES (7, 1, TY1//' Correct by inserting the '//
     &                     'following PROTRAN line: '//
     &                     '$OPTIONS;CHWORKSPACE=%(I1)',NFLAG)
               IF (NFLAG.NE.0) GOTO 9999
               NFLAGBIS=164
            ELSE
               CALL E1MES (5, 5, TY1//'Correct by calling IWKCIN '//
     &                     'from the main program with the three '//
     &                     'following statements: %/      COMMON '//
     &                     '/WKSPCH/ CHWKSP%/      CHARACTER '//
     &                     'CHWKSP(%(I1))%/      CALL '//
     &                     'IWKCIN(%(I1),%(I2))',NFLAG)
               IF (NFLAG.NE.0) GOTO 9999
               NFLAGBIS=165
            END IF
         ELSE IF (LNEED.EQ.0 .AND. LNEEDA.GT.0) THEN
            IF (XXLINE(PROLVL).GE.1 .AND. XXLINE(PROLVL).LE.999) THEN
               CALL E1MES (7, 2, TY2//' Correct by inserting the '//
     &                     'following PROTRAN line: '//
     &                     '$OPTIONS;CHWORKSPACE=%(I1),%(I2)',NFLAG)
               IF (NFLAG.NE.0) GOTO 9999
               NFLAGBIS=166
            ELSE
               CALL E1MES (5, 5, TY2//'Correct by calling IWKCIN '//
     &                     'from the main program with the three '//
     &                     'following statements: %/      COMMON '//
     &                     '/WKSPCH/ CHWKSP%/      CHARACTER '//
     &                     'CHWKSP(%(I1))%/      CALL '//
     &                     'IWKCIN(%(I1),%(I2))',NFLAG)
              IF (NFLAG.NE.0) GOTO 9999
              NFLAGBIS=167
            END IF
         ELSE
            IF (XXLINE(PROLVL).GE.1 .AND. XXLINE(PROLVL).LE.999) THEN
               CALL E1MES (7, 3, TY3//'there is insufficient '//
     &                     'workspace to  hold the allocations.  '//
     &                     'Correct by inserting the following '//
     &                     'PROTRAN line: '//
     &                     '$OPTIONS;CHWORKSPACE=%(I1),%(I2)',NFLAG)
               IF (NFLAG.NE.0) GOTO 9999
               NFLAGBIS=168
            ELSE
               CALL E1MES (5, 5, TY3//'insufficient workspace to '//
     &                     'hold the allocations. To correct, call '//
     &                     'IWKCIN from the main program thusly:%/'//
     &                     '      COMMON /WKSPCH/ CHWKSP%/'//
     &                     '      CHARACTER CHWKSP(%(I1))%/'//
     &                     '      CALL IWKCIN(%(I1),%(I2))',NFLAG)
              IF (NFLAG.NE.0) GOTO 9999
              NFLAGBIS=169
            END IF
         END IF
         CALL E1POP ('I1KCGT',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POS (5, IPA, ISA,NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1POS (7, IPA7, ISA7,NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAG=NFLAGBIS
      END IF
C
9999  RETURN
      END
