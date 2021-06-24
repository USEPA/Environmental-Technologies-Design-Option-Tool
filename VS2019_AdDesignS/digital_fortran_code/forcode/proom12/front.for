
      program front
      IMPLICIT NONE
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
c------ MAXIMUMS
      INTEGER MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE
      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=6000,MAXDE=750)
c new maximums:
      INTEGER MXTBACK
      PARAMETER (MXTBACK=400)

c------ MISCELLANEOUS
      DOUBLE PRECISION EPS_ERROR_CRITERIA
      PARAMETER (EPS_ERROR_CRITERIA=0.0005)

c------ PARAMETERS TO PSDM() SUBROUTINE
      INTEGER*2 Numb
      DOUBLE PRECISION Chemicals(MXCOMP,16)
      DOUBLE PRECISION Ads_Prop(4)
      DOUBLE PRECISION C_Prop(3)
      DOUBLE PRECISION T(MAXPTS,2)
      DOUBLE PRECISION CPVB(MXCOMP,MAXPTS)
      INTEGER*2 NITP
      DOUBLE PRECISION TT(5)
      INTEGER*2 NXX,MXX,NinI
      DOUBLE PRECISION TinI(MAXPTS),CinI(MXCOMP,MAXPTS)
      INTEGER*4 N_PW
      INTEGER*2 NumBed
      INTEGER*2 NFLAG
      DOUBLE PRECISION VARS1(15)
      DOUBLE PRECISION VARS2(MXCOMP,19)
      INTEGER*4 ISDBUG
      INTEGER*4 TELL_PSDM_SPECIAL_OUTPUT
c new parameters to psdm() subroutine:
      INTEGER*4 NB
      DOUBLE PRECISION TBACK(MXTBACK)
c new parameters to psdm() subroutine:
      DOUBLE PRECISION ROOM_VOL
      DOUBLE PRECISION ROOM_FLOWRATE
      DOUBLE PRECISION ROOM_C0(1:MXCOMP)
      DOUBLE PRECISION ROOM_EMIT(1:MXCOMP)
c      DOUBLE PRECISION RXN_RATE_CONSTANT(1:MXCOMP)
c      DOUBLE PRECISION RXN_PRODUCT(1:MXCOMP)
c      DOUBLE PRECISION RXN_RATIO(1:MXCOMP)
      CHARACTER*100 FN_MASSBAL_OUT
      CHARACTER*100 FN_CR_OUT
      CHARACTER*100 FN_CB_OUT
      DOUBLE PRECISION in_INITIAL_ROOM_CONC(1:MXCOMP)

c------ LOCAL VARIABLES
      INTEGER I,J,K
      CHARACTER*200 DUMMY1
      CHARACTER*200 DUMMY_LINE
      INTEGER BEDSIMTYPE
      INTEGER NEQ
      INTEGER stop_at_bed
      DOUBLE PRECISION EOF_TEST
      INTEGER IS_IN_ROOM
C
C------ START OF CODE.
C
C
C------ READ PATH FILE.
C
      FN_IN_FILELIST = 'PROOM1.IN'
      OPEN(UNIT=11,FILE=FN_IN_FILELIST,STATUS='OLD')
      READ(11,*) PROOM_MODE
      READ(11,*) FN_IN_MAIN
      READ(11,*) FN_OUT_SUCCESSFLAG
      READ(11,*) FN_OUT_MAIN
      READ(11,*) FN_OUT_CVST
c      READ(11,*) FN_OUT_SS_RESULTS
	CLOSE(11)
C
C------ READ IN DATA FROM INPUT FILE "INPUT.DAT"
C
c     open (unit=4,file='input.dat',status='OLD')
      open (unit=4,file=FN_IN_MAIN,status='OLD')
      read (4,*) dummy1
      read (4,*) dummy1
      read (4,*) dummy1
      read (4,*) dummy1
      read (4,*) DUMMY_LINE
      do i=1, 4
        read (4,*) ads_prop(i)
      enddo
      do i=1, 3
        read (4,*) c_prop(i)
      enddo
      read (4,*) mxx
      read (4,*) nxx
      read (4,*) numb
      read (4,*) nini
      read (4,*) numbed
      read (4,*) bedsimtype
      read (4,*) isdbug
      do i=1, 3
        read (4,*) tt(i)
      enddo
c      read (4,*) NB
c      do i=1, NB
c        read (4,*) TBACK(i)
c      enddo
      NB = 0
      read (4,*) DUMMY_LINE
      read (4,*) IS_IN_ROOM
      if (IS_IN_ROOM.EQ.1) then
        READ (4,*) ROOM_VOL
        READ (4,*) ROOM_FLOWRATE
        DO I=1, NUMB
          READ (4,*) ROOM_C0(I)
          READ (4,*) ROOM_EMIT(I)
          READ (4,*) RXN_RATE_CONSTANT(I)
          READ (4,*) RXN_PRODUCT(I)
          READ (4,*) RXN_RATIO(I)
        ENDDO
        READ (4,*) FN_MASSBAL_OUT
        READ (4,*) FN_CR_OUT
        READ (4,*) FN_CB_OUT
      else
        READ (4,*) dummy1
        READ (4,*) dummy1
        DO I=1, NUMB
          READ (4,*) dummy1
          READ (4,*) dummy1
          READ (4,*) RXN_RATE_CONSTANT(I)
          READ (4,*) RXN_PRODUCT(I)
          READ (4,*) RXN_RATIO(I)
        ENDDO      
        READ (4,*) FN_MASSBAL_OUT
        READ (4,*) FN_CR_OUT
        READ (4,*) FN_CB_OUT
      endif
      read (4,*) DUMMY_LINE
	do i=1, numb
        read (4,*) dummy1
        do j=1, 16
          read (4,*) chemicals(i,j)
        enddo
        read (4,*) in_INITIAL_ROOM_CONC(i)
      enddo
      if (nini.gt.0) then
        read (4,*) DUMMY_LINE
        read (4,*) dummy1
        read (4,*) dummy1
        do j=1,nini
          read (4,*) tini(j), (cini(i,j),i=1,numb)
        enddo
      endif
      do i=1,numb
        read (4,*) DUMMY_LINE
        read (4,*) bool_ROOM_COINI_ISTIMEVAR(i)
        if (bool_ROOM_COINI_ISTIMEVAR(i).eq.1) then
          read (4,*) int_ROOM_NCOINI(i)
          read (4,*) dummy1
          read (4,*) dummy1
          do j=1,int_ROOM_NCOINI(i)
            read (4,*) dbl_ROOM_TCOINI(i,j), dbl_ROOM_COINI(i,j)
          enddo
        endif
      enddo
      do i=1,numb
        read (4,*) DUMMY_LINE
        read (4,*) bool_ROOM_EMITINI_ISTIMEVAR(i)
        if (bool_ROOM_EMITINI_ISTIMEVAR(i).eq.1) then
          read (4,*) int_ROOM_NEMITINI(i)
          read (4,*) dummy1
          read (4,*) dummy1
          do j=1,int_ROOM_NEMITINI(i)
            read (4,*) dbl_ROOM_TEMITINI(i,j), dbl_ROOM_EMITINI(i,j)
          enddo
        endif
      enddo
      do i=1,numb
        read (4,*) DUMMY_LINE
        read (4,*) bool_ROOM_KINI_ISTIMEVAR(i)
        if (bool_ROOM_KINI_ISTIMEVAR(i).eq.1) then
          read (4,*) int_ROOM_NKINI(i)
          read (4,*) dummy1
          read (4,*) dummy1
          do j=1,int_ROOM_NKINI(i)
            read (4,*) dbl_ROOM_TKINI(i,j), dbl_ROOM_KINI(i,j)
          enddo
        endif
      enddo
      read (4,*) DUMMY_LINE
      read (4,*) EOF_TEST
      IF ( (DABS(EOF_TEST-12345.678)/12345.678).GT.0.001) THEN
        PRINT *, 'THE END OF FILE MARKER (`EOF_TEST`) WAS NOT DETECTED.'
        PRINT *, 'TERMINATING MODEL PROGRAM.'
        STOP
      ENDIF
      close (4)

      NEQ = NUMB*(MXX*(NXX+1)-1)
      IF (NEQ.GT.MAXDE) THEN
        PRINT *, 'YOU HAVE SPECIFIED THIS ' //
     &           'NUMBER OF DIFFERENTIAL EQUATIONS:'
        PRINT *, NEQ
        PRINT *, 'THE MAXIMUM NUMBER IS:'
        PRINT *, MAXDE
        PRINT *, 'TERMINATING MODEL PROGRAM.'
        STOP
      ENDIF
      N_PW = NEQ*NEQ + 2*NEQ

      if (bedsimtype.eq.0) then
        TELL_PSDM_SPECIAL_OUTPUT = 1
        call PSDM(Numb,Chemicals,Ads_Prop,C_Prop,
     &    T,CPVB,NITP,TT,NXX,MXX,
     &    NinI,TinI,CinI,NumBed,NFLAG,
     &    VARS1,VARS2,ISDBUG,
     &    TELL_PSDM_SPECIAL_OUTPUT,NB,TBACK,
     &    IS_IN_ROOM,ROOM_VOL,ROOM_FLOWRATE,ROOM_C0,
     &    ROOM_EMIT,
     &    FN_MASSBAL_OUT,FN_CR_OUT,FN_CB_OUT,
     &    in_INITIAL_ROOM_CONC)
c     &    RXN_RATE_CONSTANT,RXN_PRODUCT,RXN_RATIO,
	endif
      if (bedsimtype.eq.1) then
C
C     Recalculation of length (L) and weight (WT) based on the
C     number of axial elements.
C
        ADS_PROP(1) = ADS_PROP(1)/DBLE(NUMBED)
        ADS_PROP(3) = ADS_PROP(3)/DBLE(NUMBED)
C
C     Changing time of first output point to very small value if
C     more than one axial element is desired.
C
        IF (NUMBED .GT. 1) THEN
          TT(2) = 1.0d-8
        ENDIF

        stop_at_bed = numbed
        do i=1, stop_at_bed
          PRINT *,'PERFORMING CALCULATIONS FOR AXIAL ELEMENT ', I
          numbed = i
          if (i.eq.stop_at_bed) then
            TELL_PSDM_SPECIAL_OUTPUT = 1
          else
            TELL_PSDM_SPECIAL_OUTPUT = 0
          endif
          call PSDM(Numb,Chemicals,Ads_Prop,C_Prop,
     &      T,CPVB,NITP,TT,NXX,MXX,
     &      NinI,TinI,CinI,NumBed,NFLAG,
     &      VARS1,VARS2,ISDBUG,
     &      TELL_PSDM_SPECIAL_OUTPUT,NB,TBACK,
     &      IS_IN_ROOM,ROOM_VOL,ROOM_FLOWRATE,ROOM_C0,
     &      ROOM_EMIT,
     &      FN_MASSBAL_OUT,FN_CR_OUT,FN_CB_OUT,
     &      in_INITIAL_ROOM_CONC)
c     &    RXN_RATE_CONSTANT,RXN_PRODUCT,RXN_RATIO,
          NINI = NITP
          do J=1, NITP
            TINI(J) = T(J,1)
            do K=1, NUMB
              if (CPVB(K,J).LT.EPS_ERROR_CRITERIA) then
                CINI(K,J) = EPS_ERROR_CRITERIA
              else
C------- CONVERT FROM DIMENSIONLESS (C/C0) TO UG/L
                CINI(K,J) = CPVB(K,J)*CHEMICALS(K,2)
              endif
            enddo
          enddo          
        enddo
      endif      

      CALL GENERATE_NFLAG_OUTPUT(NFLAG,0.0D0)


c: Still need to work on:

c           cpvb        O
c           nitp        O  X
c           nini        i  x
c           tini        i  x
c           cini        i  x
c           n_pw        i  x
c           nflag       O
           


      end



CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
CC
CC    SUBROUTINE GENERATE_NFLAG_OUTPUT
CC
CC    PURPOSE: PERFORM A FEW FINAL OUTPUTS.
CC
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      SUBROUTINE GENERATE_NFLAG_OUTPUT(NFLAG_DGEAR_SIMULATION,USE_TOUT)
      IMPLICIT NONE
      INTEGER*2 NFLAG_DGEAR_SIMULATION
	DOUBLE PRECISION USE_TOUT
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
C
C------ LOCAL VARIABLES.
C
	INTEGER NFLAG
      INTEGER I

C1010  FORMAT(/
C     &  'NFLAG_DGEAR_SIMULATION, ERROR FLAG.............. = ',I5/
C     &  )

      NFLAG = NFLAG_DGEAR_SIMULATION
c      DO I=1,2
      DO I=1,1
	  IF (I.EQ.1) THEN
C
C-------- ERASE ORIGINAL FILE IF EXISTS.
C
          OPEN(UNIT=15,FILE=FN_OUT_SUCCESSFLAG)
        ENDIF
c	  IF (I.EQ.2) THEN
cC
cC-------- APPEND TO ORIGINAL FILE.
cC
c          OPEN(UNIT=15,FILE=FN_OUT_MAIN,ACCESS='APPEND')
c        ENDIF
        WRITE(15,*) 'NFLAG_DGEAR_SIMULATION, ' //
     &              'ERROR FLAG FOR SIMULATION:'
        WRITE(15,*) NFLAG_DGEAR_SIMULATION
c        IF (NFLAG.NE.0) THEN
c          WRITE(15,*) 'ERROR OCCURRED AT T (MINUTES) = ', 
c     &                USE_TOUT/TCONV
c	  ENDIF
        WRITE(15,*) 'DESCRIPTION OF THIS ERROR:'
        IF (NFLAG .EQ. 15) THEN
          WRITE(15,2015)
        ELSE IF (NFLAG .EQ. 105) THEN
	    WRITE(15,2105) 
        ELSE IF (NFLAG .EQ. 115) THEN
	    WRITE(15,2115)
        ELSE IF (NFLAG .EQ. 155) THEN
	    WRITE(15,2155)
        ELSE IF (NFLAG .EQ. 205) THEN
	    WRITE(15,2205)
        ELSE IF (NFLAG .EQ. 255) THEN
	    WRITE(15,2255)
        ELSE IF (NFLAG .EQ. 305) THEN
	    WRITE(15,2305)
        ELSE IF (NFLAG .EQ. 405) THEN
	    WRITE(15,2405)
        ELSE IF (NFLAG .EQ. 415) THEN
	    WRITE(15,2415)
        ELSE IF (NFLAG .EQ. 425) THEN
	    WRITE(15,2425)
        ELSE IF (NFLAG .EQ. 435) THEN
	    WRITE(15,2435)
        ELSE IF (NFLAG .EQ. 445) THEN
	    WRITE(15,2445)
        ELSE IF (NFLAG .EQ. 1603) THEN
	    WRITE(15,*) 'RUN ABORTED SINCE NOT ENOUGH ' //
     &      'WORKSPACE AVAILABLE' //
     &      'FOR DATA STORAGE, CHANGE ALLOCATION'
        ELSE IF (NFLAG .EQ. 1901) THEN
	    WRITE(15,*)'RUN ABORTED DUE TO SOME ' //
     &      'TYPE OF USER MIS-INPUT ' //
     &      'OR UN-DOCUMENTED INTERNAL ERROR.'
        ENDIF
        CLOSE(15)
      ENDDO

2015  FORMAT(1X,'WARNING..  T + H = T ON NEXT STEP.')   
2105  FORMAT(1X,//,'KFLAG = -1 FROM INTEGRATOR, ERROR TEST FAILED',/)
2115  FORMAT(1X,' H HAS BEEN REDUCED AND STEP WILL BE RETRIED',//)
2155  FORMAT(//44H PROBLEM APPEARS UNSOLVABLE WITH GIVEN INPUT//)  

c--- New code modified by ejoman on 1999-May-11 begins:
2205  FORMAT(//35H 'KFLAG = -2 FROM INTEGRATOR' 
     1	/52H 'THE REQUESTED ERROR IS SMALLER THAN CAN BE HANDLED'//)
c--- New code ends.
c--- Old code begins:
c2205  FORMAT(//35H KFLAG = -2 FROM INTEGRATOR 
c     1	/52H  THE REQUESTED ERROR IS SMALLER THAN CAN BE HANDLED//) 
c--- Old code ends.

2255  FORMAT(//40H EPS TOO SMALL FOR THE MACHINE PRECISION/)
2305  FORMAT (1X,//,'CORRECTOR CONVERGENCE COULD NOT BE ACHIEVED',/)
2405  FORMAT (//28H ILLEGAL INPUT.. EPS .LE. 0.//)  
2415  FORMAT (//25H ILLEGAL INPUT.. N .LE. 0//) 
2425  FORMAT (//36H ILLEGAL INPUT.. (T0-TOUT)*H .GE. 0.//)  
2435  FORMAT (//24H ILLEGAL INPUT.. INDEX =,I5//)   
 2445 FORMAT (1X,//,'INDEX = -1 ON INPUT WITH (T-TOUT)*H .GE. 0.',/,' 
     1	INTERPOLATION WAS DONE AS ON NORMAL RETURN, DESIRED PARAMETER
     2  CHANGES WERE NOT MADE.') 

      RETURN
	END



C
C                         FUNCTION INTERP_CO(I,T)
C                                                                       
C   This function returns the value of the room influent concentration
C   for chemical "I" at time "T". Interpolation is performed
C   if necessary.
C                                                                       
C   Description of Inputs/Outputs/Return:
C   =====================================
C
C   *I*   I           Index of component
C   *I*   Tdim        Time at which to calculate influent conc. (minutes)
C   *R*   INTERP_CO   Calculated value of Co (ug/L)
C

      DOUBLE PRECISION FUNCTION INTERP_CO(I,Tdim)
      IMPLICIT NONE                              
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
cc new parameters to psdm() subroutine:
c      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
c      INTEGER IS_IN_ROOM
c      DOUBLE PRECISION ROOM_VOL
c      DOUBLE PRECISION ROOM_FLOWRATE
c      DOUBLE PRECISION ROOM_C0(1:MXCOMP)
c      DOUBLE PRECISION ROOM_EMIT(1:MXCOMP)
c      CHARACTER*100 FN_MASSBAL_OUT
cc new parameters to psdm() subroutine:
c      COMMON /AMWAY1/ IS_IN_ROOM,ROOM_VOL,ROOM_FLOWRATE,ROOM_C0,
c     &  ROOM_EMIT,FN_MASSBAL_OUT
c------ LOCAL VARIABLES
      INTEGER I
      INTEGER J
      DOUBLE PRECISION Tdim
      INTEGER NMAX
C
C------ START OF CODE.
C
      NMAX = int_ROOM_NCOINI(I)
      IF (bool_ROOM_COINI_ISTIMEVAR(I).EQ.0) THEN
        INTERP_CO = 0.0D0
c        INTERP_CO = ROOM_C0(I)
        GOTO 99
      ENDIF
      IF (NMAX.LE.0) THEN
        INTERP_CO = 0.0D0
c        INTERP_CO = ROOM_C0(I)
        GOTO 99
      ENDIF
      IF (Tdim .LE. dbl_ROOM_TCOINI(I,1) ) THEN                                          
	  INTERP_CO = dbl_ROOM_COINI(I,1)
      ELSE IF (Tdim .GE. dbl_ROOM_TCOINI(I,NMAX) ) THEN                                   
	  INTERP_CO = dbl_ROOM_COINI(I,NMAX)                                              
      ELSE                                                              
	  J = 1                                                          
   10   J = J + 1                                                      
	  IF(Tdim .GE. dbl_ROOM_TCOINI(I,J-1) .AND. 
     &     Tdim .LE. dbl_ROOM_TCOINI(I,J) ) THEN               
	    INTERP_CO = 
     &        dbl_ROOM_COINI(I,J-1) + 
     &        (dbl_ROOM_COINI(I,J)-dbl_ROOM_COINI(I,J-1))*
     &        (Tdim-dbl_ROOM_TCOINI(I,J-1))/  
     &        (dbl_ROOM_TCOINI(I,J)-dbl_ROOM_TCOINI(I,J-1))                                     
	  ELSE IF (J .LT. NMAX ) THEN                                  
          GOTO 10                                                 
	  ENDIF                                                       
      ENDIF                                                             
99    RETURN                                                            
      END    


C
C                         FUNCTION INTERP_WA(I,T)
C                                                                       
C   This function returns the value of the room mass emission rate
C   for chemical "I" at time "T". Interpolation is performed
C   if necessary.
C                                                                       
C   Description of Inputs/Outputs/Return:
C   =====================================
C
C   *I*   I           Index of component
C   *I*   Tdim        Time at which to calculate influent conc. (minutes)
C   *R*   INTERP_WA   Calculated value of w*A (ug/s)
C

      DOUBLE PRECISION FUNCTION INTERP_WA(I,Tdim)
      IMPLICIT NONE                              
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
cc new parameters to psdm() subroutine:
c      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
c      INTEGER IS_IN_ROOM
c      DOUBLE PRECISION ROOM_VOL
c      DOUBLE PRECISION ROOM_FLOWRATE
c      DOUBLE PRECISION ROOM_C0(1:MXCOMP)
c      DOUBLE PRECISION ROOM_EMIT(1:MXCOMP)
c      CHARACTER*100 FN_MASSBAL_OUT
cc new parameters to psdm() subroutine:
c      COMMON /AMWAY1/ IS_IN_ROOM,ROOM_VOL,ROOM_FLOWRATE,ROOM_C0,
c     &  ROOM_EMIT,FN_MASSBAL_OUT
c------ LOCAL VARIABLES
      INTEGER I
      INTEGER J
      DOUBLE PRECISION Tdim
      INTEGER NMAX
C
C------ START OF CODE.
C
      NMAX = int_ROOM_NEMITINI(I)
      IF (bool_ROOM_EMITINI_ISTIMEVAR(I).EQ.0) THEN
        INTERP_WA = 0.0D0
c        INTERP_WA = ROOM_EMIT(I)
        GOTO 99
      ENDIF
      IF (NMAX.LE.0) THEN
        INTERP_WA = 0.0D0
c        INTERP_WA = ROOM_EMIT(I)
        GOTO 99
      ENDIF
      IF (Tdim .LE. dbl_ROOM_TEMITINI(I,1) ) THEN                                          
	  INTERP_WA = dbl_ROOM_EMITINI(I,1)
      ELSE IF (Tdim .GE. dbl_ROOM_TEMITINI(I,NMAX) ) THEN                                   
	  INTERP_WA = dbl_ROOM_EMITINI(I,NMAX)                                              
      ELSE                                                              
	  J = 1                                                          
   10   J = J + 1                                                      
	  IF(Tdim .GE. dbl_ROOM_TEMITINI(I,J-1) .AND. 
     &     Tdim .LE. dbl_ROOM_TEMITINI(I,J) ) THEN               

c      print *, '--',Tdim
c      print *, '       ',
c     &        dbl_ROOM_EMITINI(I,J-1) + 
c     &        (dbl_ROOM_EMITINI(I,J)-dbl_ROOM_EMITINI(I,J-1))*
c     &        (Tdim-dbl_ROOM_TEMITINI(I,J-1))/  
c     &        (dbl_ROOM_TEMITINI(I,J)-dbl_ROOM_TEMITINI(I,J-1))                                     
	    INTERP_WA = 
     &        dbl_ROOM_EMITINI(I,J-1) + 
     &        (dbl_ROOM_EMITINI(I,J)-dbl_ROOM_EMITINI(I,J-1))*
     &        (Tdim-dbl_ROOM_TEMITINI(I,J-1))/  
     &        (dbl_ROOM_TEMITINI(I,J)-dbl_ROOM_TEMITINI(I,J-1))                                     
	  ELSE IF (J .LT. NMAX ) THEN                                  
          GOTO 10                                                 
	  ENDIF                                                       
      ENDIF                                                             
99    RETURN                                                            
      END    


C
C                         FUNCTION INTERP_K(I,T)
C                                                                       
C   This function returns the value of the Freundlich K
C   for chemical "I" at time "T". Interpolation is performed
C   if necessary.
C                                                                       
C   Description of Inputs/Outputs/Return:
C   =====================================
C
C   *I*   I           Index of component
C   *I*   Tdim        Time at which to calculate Freundlich K (minutes)
C   *R*   INTERP_K    Calculated value of K, units of: (umol/g)*(L/umol)^(1/n)
C

      DOUBLE PRECISION FUNCTION INTERP_K(I,Tdim)
      IMPLICIT NONE                              
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
cc new parameters to psdm() subroutine:
c      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
c      INTEGER IS_IN_ROOM
c      DOUBLE PRECISION ROOM_VOL
c      DOUBLE PRECISION ROOM_FLOWRATE
c      DOUBLE PRECISION ROOM_C0(1:MXCOMP)
c      DOUBLE PRECISION ROOM_EMIT(1:MXCOMP)
c      CHARACTER*100 FN_MASSBAL_OUT
cc new parameters to psdm() subroutine:
c      COMMON /AMWAY1/ IS_IN_ROOM,ROOM_VOL,ROOM_FLOWRATE,ROOM_C0,
c     &  ROOM_EMIT,FN_MASSBAL_OUT
c------ LOCAL VARIABLES
      INTEGER I
      INTEGER J
      DOUBLE PRECISION Tdim
      INTEGER NMAX
C
C------ START OF CODE.
C
      NMAX = int_ROOM_NKINI(I)
      IF (bool_ROOM_KINI_ISTIMEVAR(I).EQ.0) THEN
        INTERP_K = 0.0D0
c        INTERP_K = ROOM_K(I)
        GOTO 99
      ENDIF
      IF (NMAX.LE.0) THEN
        INTERP_K = 0.0D0
c        INTERP_K = ROOM_K(I)
        GOTO 99
      ENDIF
      IF (Tdim .LE. dbl_ROOM_TKINI(I,1) ) THEN                                          
	  INTERP_K = dbl_ROOM_KINI(I,1)
      ELSE IF (Tdim .GE. dbl_ROOM_TKINI(I,NMAX) ) THEN                                   
	  INTERP_K = dbl_ROOM_KINI(I,NMAX)                                              
      ELSE                                                              
	  J = 1                                                          
   10   J = J + 1                                                      
	  IF(Tdim .GE. dbl_ROOM_TKINI(I,J-1) .AND. 
     &     Tdim .LE. dbl_ROOM_TKINI(I,J) ) THEN               

c      print *, '--',Tdim
c      print *, '       ',
c     &        dbl_ROOM_KINI(I,J-1) + 
c     &        (dbl_ROOM_KINI(I,J)-dbl_ROOM_KINI(I,J-1))*
c     &        (Tdim-dbl_ROOM_TKINI(I,J-1))/  
c     &        (dbl_ROOM_TKINI(I,J)-dbl_ROOM_TKINI(I,J-1))                                     
	    INTERP_K = 
     &        dbl_ROOM_KINI(I,J-1) + 
     &        (dbl_ROOM_KINI(I,J)-dbl_ROOM_KINI(I,J-1))*
     &        (Tdim-dbl_ROOM_TKINI(I,J-1))/  
     &        (dbl_ROOM_TKINI(I,J)-dbl_ROOM_TKINI(I,J-1))                                     
	  ELSE IF (J .LT. NMAX ) THEN                                  
          GOTO 10                                                 
	  ENDIF                                                       
      ENDIF                                                             
99    RETURN                                                            
      END    

