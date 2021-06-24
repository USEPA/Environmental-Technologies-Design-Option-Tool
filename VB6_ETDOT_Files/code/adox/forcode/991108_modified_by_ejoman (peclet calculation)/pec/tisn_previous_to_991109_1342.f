      PROGRAM TANKNUMBER
	
      IMPLICIT NONE
     
	INTEGER ::I
	INTEGER,PARAMETER :: S=1600
	REAL,PARAMETER :: PI=3.14159265359,TOL=1E-8
      INTEGER:: NDATA,FLAG,NIT
      REAL::CO,TBAR,CT,VAR,INTV,CNOT,PEC,PEO,
	+    ALPHA,Q1,Q2,X1,X2
	INTEGER ::NTANK
      
      REAL,DIMENSION(1:S)::ETHETA1,ETHETA2,T,C,
     +	 THETA,MURT
c////////////////////////////////////////////////////////////////////////////////////////
c---- Code by Eric J. Oman (1999-Sep-29) begins.
      REAL,DIMENSION(1:S)::OUTPT2_THETA,OUTPT2_E
      INTEGER OUTPT2_COUNT
      REAL OUTPT2_LOW,OUTPT2_HIGH,OUTPT2_THETA_STEP
c---- Code by Eric J. Oman (1999-Sep-29) ends.
c////////////////////////////////////////////////////////////////////////////////////////
c////////////////////////////////////////////////////////////////////////////////////////
c---- Code by Eric J. Oman (1999-Oct-20) begins.
      REAL,DIMENSION(1:S)::OUTPT3_THETA,OUTPT3_E
      INTEGER OUTPT3_COUNT
c---- Code by Eric J. Oman (1999-Oct-20) ends.
c////////////////////////////////////////////////////////////////////////////////////////
	

C  This program takes the data from a pulse dye study of a real system 
C  and calculates the exit age distribution. It then uses this data to 
C  determine how many tanks should used to simulate the real system.
C
C  PROGRAMMED BY :
C                    Ke Li
C                    Department of Civil & Environmental Engineering
C                    Michigan Technological University
C
C  NOMENCLATURE
C    INPUT	 - input file of the form NDATA/T(i), C(i)
C    OUTPUT	 - output file for results of all calculations
C    NDATA	 - number of data points of experimental data
C    T(i)	 - experimental time data (maximum length = 1600)
C    C(i)	 - experimental concentration data (maximum length = 1600)
C    CO	 - initial concentration
C    TBAR	 - mean residence time
C    THETA	 - dimensionless time
C    ETHETA1(i) = E(theta) for experimental data
C    ETHETA2(i) = predicted E(theta) for NTIS model


      OPEN(UNIT=7, FILE="input.txt", STATUS="unknown")
      OPEN(UNIT=8, FILE="outpt.txt", STATUS="unknown")

      READ(7,*) NDATA
	DO 10 I=1, NDATA
         READ(7,*) T(I), C(I)
10    CONTINUE

      ENDFILE (UNIT=7)
      CLOSE (UNIT=7)

C  Calculate TBAR and CO using trapezoidal integration method

      CNOT = 0.0
      CT = 0.0

      DO 20 I=1, NDATA-1
         CNOT = CNOT + (C(I)+C(I+1))*(T(I+1)-T(I))/2.0
         CT = CT + (T(I+1)*C(I+1)+T(I)*C(I))*(T(I+1)-T(I))/2.0
20    CONTINUE

      TBAR = CT/CNOT
      CO = CNOT/TBAR

C  Convert C to dimensionless quantities

      DO I=1, NDATA
         ETHETA1(I) = C(I)/CO
         THETA(I) = T(I)/TBAR
	END DO

	CALL OBJ(NDATA,THETA,ETHETA1,TOL,ETHETA2,NTANK)

c////////////////////////////////////////////////////////////////////////////////////////
c
c---- Code by Ke Li (1999-Nov-08) begins:
c
C     these codes should be added into the main program before those 
C	WRITE statement 

C	    calculate the variance

      VAR = 0.0

      DO I=1, NDATA-1
      VAR = VAR+(((THETA(I+1)-1.0)**2)*ETHETA1(I+1)+((THETA(I)-1.0)
     +       **2)*ETHETA1(I))*(THETA(I+1)-THETA(I))/2
      END DO	

C	    calculate the Peclect number

      X1 = 1.0E-8
      X2 = 1.0E10

      PEC = RTSAFE(X1,X2,TOL)
c
c---- Code by Ke Li (1999-Nov-08) ends.
c
c////////////////////////////////////////////////////////////////////////////////////////
C      PEC = 2.0D0*DBLE(NTANK)
C      PEC = 1.0D0
      WRITE(8,2010) TBAR
      WRITE(8,2020) NTANK
      WRITE(8,2030) PEC
      
      WRITE(8,1004)
      DO 80, I=1, NDATA
         WRITE(8,1005) THETA(I), ETHETA1(I)
80    CONTINUE

      WRITE(8,1006)
      DO 90, I=1, NDATA
         WRITE(8,1005) THETA(I), ETHETA2(I)
90    CONTINUE



1004  FORMAT(//,1x,'Experimental E(theta) results : ',//,T9,'THETA',T29,
     +'E(theta)')
1005  FORMAT(T5,E12.6,T25,E12.6)
1006  FORMAT(//,1X,'Predictions of E(theta) for nTIS model:'
     +,//,T9,'THETA',T29,'E(theta)')

2010  FORMAT(1X,'Mean Residence Time (Tbar) (user-defined units) :',
     &  T40,E12.6)
2020  FORMAT(1X,'Number of tanks (dimensionless) :',T40,I6)
2030  FORMAT(1X,'Peclet Number (dimensionless)   :',T40,E12.6)

cccc	WRITE(8,*)NTANK," of tanks should be used to simulate the system."
c	WRITE(8,*) ' '
c	WRITE(8,*) ' '
c	WRITE(8,*) NTANK," tanks should be used to simulate this system."

c////////////////////////////////////////////////////////////////////////////////////////
c
c---- Code by Eric J. Oman (1999-Sep-29) begins.
c
      OPEN(UNIT=9, FILE="outpt2.txt", STATUS="unknown")
      OUTPT2_COUNT = 100
      OUTPT2_LOW = THETA(1)
      OUTPT2_HIGH = THETA(NDATA)
      OUTPT2_THETA_STEP = (OUTPT2_HIGH-OUTPT2_LOW)/DBLE(OUTPT2_COUNT-1)
      DO I=1,OUTPT2_COUNT
        OUTPT2_THETA(I) = OUTPT2_LOW + OUTPT2_THETA_STEP*DBLE(I-1)
        OUTPT2_E(I) = ETIS(NTANK,OUTPT2_THETA(I))
      ENDDO      
      WRITE(9,*) 'FILE OUTPT2.TXT'
      WRITE(9,*) 'THE FIRST LINE IS THE NUMBER OF ROWS'
      WRITE(9,*) 'EACH SUBSEQUENT ROW INCLUDES THETA AND E(THETA)'
      WRITE(9,*) '    THETA = DIMENSIONLESS TIME'
      WRITE(9,*) '    E(THETA) = PREDICTED DIMENSIONLESS CONCENTRATION'
      WRITE(9,*) '-----------------------------------------------------'
      WRITE(9,*) OUTPT2_COUNT
      DO I=1,OUTPT2_COUNT
        WRITE (9,*) OUTPT2_THETA(I), OUTPT2_E(I)
      ENDDO
      CLOSE(9)      
c
c---- Code by Eric J. Oman (1999-Sep-29) ends.
c
c////////////////////////////////////////////////////////////////////////////////////////
c////////////////////////////////////////////////////////////////////////////////////////
c
c---- Code by Eric J. Oman (1999-Oct-20) begins.
c
      OPEN(UNIT=9, FILE="outpt3.txt", STATUS="unknown")
      OUTPT3_COUNT = NDATA
      DO I=1,OUTPT3_COUNT
        OUTPT3_THETA(I) = THETA(I)
        OUTPT3_E(I) = ETHETA1(I)
      ENDDO
      WRITE(9,*) 'FILE OUTPT3.TXT'
      WRITE(9,*) 'THE FIRST LINE IS THE NUMBER OF ROWS'
      WRITE(9,*) 'EACH SUBSEQUENT ROW INCLUDES THETA AND E(THETA)'
      WRITE(9,*) '    THETA = DIMENSIONLESS TIME'
      WRITE(9,*) '    E(THETA) = EXPERIMENTAL DIMENSIONLESS CONC.'
      WRITE(9,*) '-----------------------------------------------------'
      WRITE(9,*) OUTPT3_COUNT
      DO I=1,OUTPT3_COUNT
        WRITE (9,*) OUTPT3_THETA(I), OUTPT3_E(I)
      ENDDO
      CLOSE(9)      
c
c---- Code by Eric J. Oman (1999-Oct-20) ends.
c
c////////////////////////////////////////////////////////////////////////////////////////

	CONTAINS

c////////////////////////////////////////////////////////////////////////////////////////
c
c---- Code by Ke Li (1999-Nov-08) begins:
c
C      two subroutine needed to calculate Peclect number
C	 The idea here is to find a solution of Pe by solving
C	 the function in FUNCD by bisection algorithm      
      
      SUBROUTINE FUNCD(X,F,DF)
      
      REAL:: X,F,DF
      
        F = VAR-(2/X)+(2.0/(X**2.0))*(1-EXP(-X))
        DF = 2/X**2+2*(EXP(-X))/X**2-4*(1-EXP(-X))/X**3
	              
      END SUBROUTINE FUNCD


      REAL FUNCTION rtsafe(x1,x2,xacc)
      REAL:: x1,x2,xacc
      
      INTEGER,PARAMETER::MAXIT=1000
      INTEGER j
      REAL:: df,dx,dxold,f,fh,fl,temp,xh,xl

      call funcd(x1,fl,df)						    
      call funcd(x2,fh,df)
      
	if((fl.gt.0..and.fh.gt.0.).or.(fl.lt.0..and.fh.lt.0.)) then
	pause 'root must be bracketed in rtsafe - press [Enter]'
      end if
	if(fl.eq.0.)then
        rtsafe=x1
        return
      else if(fh.eq.0.)then
        rtsafe=x2
        return
      else if(fl.lt.0.)then
        xl=x1
        xh=x2
      else
        xh=x1
        xl=x2
      endif
      
	rtsafe=.5*(x1+x2)
      dxold=abs(x2-x1)
      dx=dxold
      call funcd(rtsafe,f,df)
      do j=1,MAXIT
      if(((rtsafe-xh)*df-f)*((rtsafe-xl)*df-f).ge.0..or. abs(2.*
     +f).gt.abs(dxold*df) ) then
          dxold=dx
          dx=0.5*(xh-xl)
          rtsafe=xl+dx
          if(xl.eq.rtsafe)THEN
	    return
	    end if
        else	
          dxold=dx
          dx=f/df
          temp=rtsafe
          rtsafe=rtsafe-dx
          if(temp.eq.rtsafe)then
	    return
	    end if
        end if
        if(abs(dx).lt.xacc)then
	    return
	  end if
        call funcd(rtsafe,f,df)
        if(f.lt.0.) then
          xl=rtsafe
        else
          xh=rtsafe
        end if
	END DO
	
      pause 'rtsafe exceeding maximum iterations - press [Enter]'
      return

      END FUNCTION RTSAFE
c
c---- Code by Ke Li (1999-Nov-08) ends.
c
c////////////////////////////////////////////////////////////////////////////////////////


	REAL FUNCTION ETIS(N,TIME)
	    
	  IMPLICIT NONE
	  INTEGER,INTENT(IN)::N
	  REAL,INTENT(IN):: TIME 
	  INTEGER::I,MULTI
	  REAL::TEMP_THETA
		  
	  MULTI=1
	  TEMP_THETA=1

	  DO I=1,N-1
	    MULTI=MULTI*I
	    TEMP_THETA=TEMP_THETA*TIME
	  END DO

	  ETIS=N**N*TEMP_THETA*EXP(-N*TIME)/MULTI

	END FUNCTION ETIS
	
	REAL FUNCTION OBJFUNC(A,B,NO)
	    
	  IMPLICIT NONE
	  INTEGER::I
	  REAL,DIMENSION(1:),INTENT(IN)::A,B
	  INTEGER,INTENT(IN)::NO
	  REAL::TEMP
		    
	  TEMP=0

	  DO I=2,NO
	    TEMP=TEMP+(A(I)-B(I))**2
	  END DO

	  OBJFUNC=SQRT(TEMP/(NO-1))

	END FUNCTION OBJFUNC

	     		    
	SUBROUTINE OBJ(NO,TIME,EXP_E,TOLERANCE,ETHETATIS,TANKN)
	
        IMPLICIT NONE
	  INTEGER :: I,J,K,L
	  INTEGER,INTENT(IN)::NO
	  REAL,INTENT(IN)::TOLERANCE
	  INTEGER,INTENT(OUT)::TANKN
	  REAL,DIMENSION(1:),INTENT(IN)::TIME,EXP_E
	  REAL,DIMENSION(1:),INTENT(OUT)::ETHETATIS
	  REAL,DIMENSION(1:NO)::E_TIS
	  REAL,DIMENSION(1:1000)::ERR
	  REAL ::MIN_ERR

	  ERR(1)=0
	  MIN_ERR=1E8
	  I=2
	  
	  DO 
	    DO J=1,NO 
	      E_TIS(J)=ETIS(I,TIME(J))
	    END DO
	
	    ERR(I)=OBJFUNC(E_TIS,EXP_E,NO)

	    IF (ERR(I)<MIN_ERR) THEN
	      MIN_ERR=ERR(I)  
	      TANKN=I
	      DO K=1,NO
		 ETHETATIS(K)=E_TIS(K)
	      END DO
          END IF

	    IF (ERR(I)<TOLERANCE) THEN
	      TANKN=I
	      EXIT
	    ELSE IF (I>2) THEN
	      IF(ERR(I)>ERR(I-1).OR.ABS(ERR(I)-ERR(I-1))<TOLERANCE) THEN
		 TANKN=I-1
		 EXIT
	      END IF
	    END IF
	   
	  I=I+1

        END DO	 			 
			
	END SUBROUTINE OBJ

	END PROGRAM TANKNUMBER
