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

      WRITE(8,1008) TBAR
      
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

1008  FORMAT(1X,'Mean Residence Time (Tbar) :',T40,E12.6)

	WRITE(8,*)NTANK," of tanks should be used to simulate the system."

	CONTAINS

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
