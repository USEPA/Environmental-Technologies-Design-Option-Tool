C-----------------------------------------------------------------------
C  IMSL Name:  I1KGT
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    January 17, 1984
C
C  Purpose:    Allocate numerical workspace.
C
C  Usage:      I1KGT(NELMTS,ITYPE)
C
C  Arguments:
C     NELMTS - Number of elements of data type ITYPE to be
C              allocated.  (Input)
C     ITYPE  - Data type of array to be allocated.  (Input)
C                 1 - logical
C                 2 - integer
C                 3 - real
C                 4 - double precision
C                 5 - complex
C                 6 - double complex
C     I1KGT  - Integer function.  (Output)  Returns the index of the
C              first element in the current allocation.
C
C  Remarks:
C  1. On return, the array will occupy
C     WKSP(I1KGT), WKSP(I1KGT+1), ..., WKSP(I1KGT+NELMTS-1) where
C     WKSP is an array of data type ITYPE equivalenced to RWKSP.
C
C  2. If I1KGT is negative, the absolute value of I1KGT is the
C     additional workspace needed for the current allocation.
C
C  3. The allocator reserves the first sixteen integer locations of
C     the stack for its own internal bookkeeping.  These are initialized
C     by the function IWKIN upon the first call to the allocation
C     package.
C
C  4. The use of the first ten integer locations is as follows:
C      WKSP( 1) - LOUT    The number of current allocations
C      WKSP( 2) - LNOW    The current active length of the stack
C      WKSP( 3) - LUSED   The maximum value of WKSP(2) achieved
C                         thus far
C      WKSP( 4) - LBND    The lower bound of permanent storage which
C                         is one numeric storage unit more than the
C                         maximum allowed length of the stack.
C      WKSP( 5) - LMAX    The maximum length of the storage array
C      WKSP( 6) - LALC    The total number of allocations handled by
C                         I1KGT
C      WKSP( 7) - LNEED   The number of numeric storage units by which
C                         the array size must be increased for all past
C                         allocations to succeed
C      WKSP( 8) - LBOOK   The number of numeric storage units used for
C                         bookkeeping
C      WKSP( 9) - LCHAR   The pointer to the portion of the permanent
C                         stack which contains the bookkeeping and
C                         pointers for the character workspace
C                         allocation.
C      WKSP(10) - LLCHAR  The length of the array beginning at LCHAR
C                         set aside for character workspace bookkeeping
C                         and pointers.
C                 NOTE -  If character workspace is not being used,
C                         LCHAR and LLCHAR can be ignored.
C  5. The next six integer locations contain values describing the
C     amount of storage allocated by the allocation system to the
C     various data types.
C      WKSP(11) - Numeric storage units allocated to LOGICAL
C      WKSP(12) - Numeric storage units allocated to INTEGER
C      WKSP(13) - Numeric storage units allocated to REAL
C      WKSP(14) - Numeric storage units allocated to DOUBLE PRECISION
C      WKSP(15) - Numeric storage units allocated to COMPLEX
C      WKSP(16) - Numeric storage units allocated to DOUBLE COMPLEX
C
C  Copyright:  1984 by IMSL, Inc. All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C

      INTEGER FUNCTION I1KGT (NELMTS, ITYPE,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NELMTS, ITYPE,NFLAG,NFLAGBIS
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IDUMAL, IGAP, ILEFT, IPA, IPA7, ISA, ISA7,
     &           ISIZE(6), JTYPE, LALC, LBND, LBOOK, LMAX, LNEED,
     &           LNEED1, LNOW, LOUT, LUSED
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    FIRST
      SAVE       FIRST
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
C                                  SPECIFICATIONS FOR EQUIVALENCE
      EQUIVALENCE (LOUT, IWKSP(1))
      EQUIVALENCE (LNOW, IWKSP(2))
      EQUIVALENCE (LUSED, IWKSP(3))
      EQUIVALENCE (LBND, IWKSP(4))
      EQUIVALENCE (LMAX, IWKSP(5))
      EQUIVALENCE (LALC, IWKSP(6))
      EQUIVALENCE (LNEED, IWKSP(7))
      EQUIVALENCE (LBOOK, IWKSP(8))
      EQUIVALENCE (ISIZE(1), IWKSP(11))
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  IABS,MAX0,MOD
      INTRINSIC  IABS, MAX0, MOD
      INTEGER    IABS, MAX0, MOD
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1POS, E1PSH, E1STI, IWKIN
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KQU
      INTEGER    I1KQU
C
      DATA FIRST/.TRUE./
C
      NFLAGBIS=0
      CALL E1PSH ('I1KGT ',NFLAG)
      IF (NFLAG.NE.0) GOTO 1000
C
      IF (FIRST) THEN
C                                  INITIALIZE WORKSPACE IF NEEDED
         FIRST = .FALSE.
         CALL IWKIN (0,NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) CALL E1POP('I1KGT ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000 
         IF (NFLAGBIS.NE.0) GOTO 1000         
       END IF
C                                  NUMBER OF ELEMENTS LESS THAN 0
      IF (NELMTS .LT. 0) THEN
         CALL E1STI (1, NELMTS)
         CALL E1MES (5, 2, 'Number of elements is not positive.%/'//
     &               'NELMTS = %(I1).',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0        
         CALL E1POP ('I1KGT ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         NFLAGBIS=4
         GO TO 9000
      END IF
C                                  ILLEGAL DATA TYPE REQUESTED
      IF (ITYPE.EQ.0 .OR. IABS(ITYPE).GE.7) THEN
         CALL E1MES (5, 3, 'Illegal data type requested.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0       
         CALL E1POP ('I1KGT ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         NFLAGBIS=5         
         GO TO 9000
      END IF
C                                  BOOKKEEPING OVERWRITTEN
      IF (LNOW.LT.LBOOK .OR. LNOW.GT.LUSED .OR. LUSED.GT.LMAX .OR.
     &    LNOW.GE.LBND .OR. LOUT.GT.LALC) THEN
         CALL E1MES (5, 4, 'One or more of the first eight '//
     &               'bookkeeping locations in IWKSP have been '//
     &               'overwritten.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0        
         
         CALL E1POP ('I1KGT ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         if (NFLAGBIS.NE.0) GOTO 1000
         NFLAGBIS=6
         GO TO 9000
      END IF
C
      CALL E1POP ('I1KGT ',NFLAG)
      IF (NFLAG.NE.0) GOTO 1000
C                                  DETERMINE NUMBER OF LOCATIONS STILL
C                                  AVAILABLE FOR DATA TYPE ITYPE
C                                  NOTE: I1KQU ALLOWS FOR 2 INTEGER
C                                        POINTERS WHICH MUST BE HANDLED
C                                        ARTIFICIALLY IF ILEFT = 0.
      ILEFT = I1KQU(IABS(ITYPE),NFLAG)
      IF (NFLAG.NE.0) GOTO 1000
C
      IF (ITYPE .GT. 0) THEN
C                                  RELEASABLE STORAGE
         IF (ILEFT .GE. NELMTS) THEN
            I1KGT = (LNOW*ISIZE(2)-1)/ISIZE(ITYPE) + 2
            I = ((I1KGT-1+NELMTS)*ISIZE(ITYPE)-1)/ISIZE(2) + 3
C                                  IWKSP(I-1) CONTAINS THE DATA TYPE FOR
C                                  THIS ALLOCATION. IWKSP(I) CONTAINS
C                                  LNOW FOR THE PREVIOUS ALLOCATION.
            IWKSP(I-1) = ITYPE
            IWKSP(I) = LNOW
            LOUT = LOUT + 1
            LALC = LALC + 1
            LNOW = I
            LUSED = MAX0(LUSED,LNOW)
            LNEED = 0
         ELSE
C                                  RELEASABLE STORAGE WAS REQUESTED
C                                  BUT THE STACK WOULD OVERFLOW.
C                                  THEREFORE, ALLOCATE RELEASABLE
C                                  SPACE THROUGH THE END OF THE STACK
            IF (LNEED .EQ. 0) THEN
               IDUMAL = (LNOW*ISIZE(2)-1)/ISIZE(ITYPE) + 2
               I = ((IDUMAL-1+ILEFT)*ISIZE(ITYPE)-1)/ISIZE(2) + 3
C                                  ADVANCE COUNTERS AND STORE POINTERS
C                                  IF THERE IS ROOM TO DO SO
               IF (I .LT. LBND) THEN
C                                  IWKSP(I-1) CONTAINS THE DATA TYPE FOR
C                                  THIS ALLOCATION. IWKSP(I) CONTAINS
C                                  LNOW FOR THE PREVIOUS ALLOCATION.
                  IWKSP(I-1) = ITYPE
                  IWKSP(I) = LNOW
                  LOUT = LOUT + 1
                  LALC = LALC + 1
                  LNOW = I
                  LUSED = MAX0(LUSED,LNOW)
               END IF
            END IF
C                                  CALCULATE AMOUNT NEEDED TO ACCOMODATE
C                                  THIS ALLOCATION REQUEST
            LNEED1 = (NELMTS-ILEFT)*ISIZE(ITYPE)
            IF (ILEFT .EQ. 0) THEN
               IGAP = ISIZE(ITYPE) - MOD(LNOW+LNEED,ISIZE(ITYPE))
               IF (IGAP .EQ. ISIZE(ITYPE)) IGAP = 0
               LNEED1 = LNEED1 + 2*ISIZE(2) + IGAP
            END IF
C                                  MODIFY LNEED ACCORDING TO THE SIZE
C                                  OF THE BASE BEING USED (D.P. HERE)
            LNEED = LNEED + ((LNEED1+ISIZE(3)-1)/ISIZE(3))
C                                  SINCE CURRENT ALLOCATION IS ILLEGAL,
C                                  RETURN THE NEGATIVE OF THE ADDITIONAL
C                                  AMOUNT NEEDED TO MAKE IT LEGAL
            I1KGT = -LNEED
         END IF
      ELSE
C                                  PERMANENT STORAGE
         IF (ILEFT .GE. NELMTS) THEN
            JTYPE = -ITYPE
            I1KGT = (LBND*ISIZE(2)-1)/ISIZE(JTYPE) + 1 - NELMTS
            I = ((I1KGT-1)*ISIZE(JTYPE))/ISIZE(2) - 1
C                                  IWKSP(I) CONTAINS LBND FOR PREVIOUS
C                                  PERMANENT STORAGE ALLOCATION.
C                                  IWKSP(I+1) CONTAINS THE DATA TYPE FOR
C                                  THIS ALLOCATION.
            IWKSP(I) = LBND
            IWKSP(I+1) = JTYPE
            LALC = LALC + 1
            LBND = I
            LNEED = 0
         ELSE
C                                  PERMANENT STORAGE WAS REQUESTED
C                                  BUT THE STACK WOULD OVERFLOW,
C                                  THEREFORE, ALLOCATE RELEASABLE
C                                  SPACE THROUGH THE END OF THE STACK
            IF (LNEED .EQ. 0) THEN
               JTYPE = -ITYPE
               IDUMAL = (LNOW*ISIZE(2)-1)/ISIZE(JTYPE) + 2
               I = ((IDUMAL-1+ILEFT)*ISIZE(JTYPE)-1)/ISIZE(2) + 3
C                                  ADVANCE COUNTERS AND STORE POINTERS
C                                  IF THERE IS ROOM TO DO SO
               IF (I .LT. LBND) THEN
C                                  IWKSP(I-1) CONTAINS THE DATA TYPE FOR
C                                  THIS ALLOCATION. IWKSP(I) CONTAINS
C                                  LNOW FOR THE PREVIOUS ALLOCATION.
                  IWKSP(I-1) = JTYPE
                  IWKSP(I) = LNOW
                  LOUT = LOUT + 1
                  LALC = LALC + 1
                  LNOW = I
                  LUSED = MAX0(LUSED,LNOW)
               END IF
            END IF
C                                  CALCULATE AMOUNT NEEDED TO ACCOMODATE
C                                  THIS ALLOCATION REQUEST
            LNEED1 = (NELMTS-ILEFT)*ISIZE(-ITYPE)
            IF (ILEFT .EQ. 0) THEN
               IGAP = ISIZE(-ITYPE) - MOD(LNOW+LNEED,ISIZE(-ITYPE))
               IF (IGAP .EQ. ISIZE(-ITYPE)) IGAP = 0
               LNEED1 = LNEED1 + 2*ISIZE(2) + IGAP
            END IF
C                                  MODIFY LNEED ACCORDING TO THE SIZE
C                                  OF THE BASE BEING USED (D.P. HERE)
            LNEED = LNEED + ((LNEED1+ISIZE(3)-1)/ISIZE(3))
C                                  SINCE CURRENT ALLOCATION IS ILLEGAL,
C                                  RETURN THE NEGATIVE OF THE ADDITIONAL
C                                  AMOUNT NEEDED TO MAKE IT LEGAL
            I1KGT = -LNEED
         END IF
      END IF
C                                  STACK OVERFLOW - UNRECOVERABLE ERROR
 9000 IF (LNEED .GT. 0) THEN
         CALL E1POS (-5, IPA, ISA,NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1POS (5, 0, 0,NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1POS (-7, IPA7, ISA7,NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1POS (7, 0, 0,NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1PSH ('I1KGT ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1STI (1, LNEED+(LMAX/ISIZE(3)))
         IF (XXLINE(PROLVL).GE.1 .AND. XXLINE(PROLVL).LE.999) THEN
            CALL E1MES (7, 1, 'Insufficient workspace for current '//
     &                  'allocation(s).  Correct by inserting the '//
     &                  'following PROTRAN line: $OPTIONS;WORKSPACE=%'//
     &                  '(I1)',NFLAG)
           IF (NFLAG.NE.0) GOTO 1000
           NFLAG=7
           GOTO 1000
         ELSE
            CALL E1MES (5, 5, 'Insufficient workspace for current '//
     &                  'allocation(s). Correct by calling IWKIN '//
     &                  'from main program with the three following '//
     &                  'statements:  (REGARDLESS OF PRECISION)%/'//
     &                  '      COMMON /WORKSP/  RWKSP%/      REAL '//
     &                  'RWKSP(%(I1))%/      CALL IWKIN(%(I1))',NFLAG)
            IF (NFLAG.NE.0) GOTO 1000
            NFLAG=8
            GOTO 1000
         END IF
         CALL E1POP ('I1KGT ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1POS (5, IPA, ISA,NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         CALL E1POS (7, IPA7, ISA7,NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
      END IF
C
      
      
1000  IF (NFLAG.EQ.0) THEN 
       NFLAG=NFLAGBIS 
      ENDIF
      RETURN
      END
