C    This is the last version of Peclect calculation and tank
C	number estimation
C
C  This program takes the data from a pulse dye study of a real system 
C  and calculates the exit age distribution. It then calculate the peclet
C  number of the system. 
C
C  PROGRAMMED BY :
C                    Ke Li
C                    Department of Civil & Environmental Engineering
C                    Michigan Technological University
C
C  NOMENCLATURE
C    pecinput - input file of the form NDATA/T(i), C(i)
C    pecoutpt - output file for results of all calculations
C    NDATA	 - number of data points of experimental data
C    T(i)	 - experimental time data (maximum length = 1600)
C    C(i)	 - experimental concentration data (maximum length = 1600)
C    CO	 - initial concentration
C    TBAR	 - mean residence time
C    VAR      - varience of exit age distribution
C    THETA	 - dimensionless time
C    ETHETA1() - E(theta) for experimental data
C    ETHETA2() - predicted E(theta) for closed system
C    ETHETA3() - predicted E(theta) for open system
C    LEFT	 - lower range of estimated peclet number (NOTE:should be positive)
C    RIHGT	 - upper range of estimated peclet number


	PROGRAM PECLET
      
        INTEGER S
        PARAMETER(S=1600,PI=3.14159265359)

        INTEGER NDATA,FLAG,NSTEPS,ITER1,ITER2
        DOUBLE PRECISION CO,TBAR,ETHETA1,ETHETA2,ETHETA3,T,C,CT,THETA,
     +    VAR,CNOT,PEC,PEO,TOLG,DFLT,X,FX,LEFT,RIGHT,PECS,PEOS,OBJO,OBJC
        DIMENSION ETHETA1(S),ETHETA2(S),ETHETA3(S),T(S),C(S),THETA(S),
     +          PECS(S),PEOS(S),OBJO(S),OBJC(S)

        EXTERNAL GOLDEN
	  COMMON/FLAG/FLAG
        COMMON/RTD/NDATA,ITER1,ITER2,ETHETA1,ETHETA2,ETHETA3,THETA,PECS,
     +           PEOS,OBJO,OBJC

c////////////////////////////////////////////////////////////////////////////////////////
c---- Code by Eric J. Oman (1999-Oct-26) begins.
      REAL,DIMENSION(1:S)::OUTPT4_THETA,OUTPT4_E
      INTEGER OUTPT4_COUNT
      REAL OUTPT4_LOW,OUTPT4_HIGH,OUTPT4_THETA_STEP
      REAL,DIMENSION(1:S)::OUTPT5_THETA,OUTPT5_E
      INTEGER OUTPT5_COUNT
      REAL OUTPT5_LOW,OUTPT5_HIGH,OUTPT5_THETA_STEP
	DOUBLE PRECISION Q1
	DOUBLE PRECISION Q2
c---- Code by Eric J. Oman (1999-Oct-26) ends.
c////////////////////////////////////////////////////////////////////////////////////////

        OPEN(UNIT=7, FILE="pecinput.txt", STATUS="unknown")
        OPEN(UNIT=8, FILE="pecoutpt.txt", STATUS="unknown")

        READ(7,*) NDATA
        DO I=1, NDATA
          READ(7,*) T(I), C(I)
	  END DO

        ENDFILE (UNIT=7)
        CLOSE (UNIT=7)

C  Calculate TBAR and CO using trapezoidal integration method

        CNOT = 0.0
        CT = 0.0

        DO I=1, NDATA-1
          CNOT = CNOT + (C(I)+C(I+1))*(T(I+1)-T(I))/2.0
          CT = CT + (T(I+1)*C(I+1)+T(I)*C(I))*(T(I+1)-T(I))/2.0
        END DO 

        TBAR = CT/CNOT
        CO = CNOT/TBAR

C  Convert C to dimensionless quantities

        DO I=1, NDATA
          ETHETA1(I) = C(I)/CO
          THETA(I) = T(I)/TBAR
	  END DO 

C  Calculation of varience (sigma**2) of the exit age distribution

        VAR = 0.0

        DO I=1, NDATA-1
          VAR = VAR+(((THETA(I+1)-1.0)**2)*ETHETA1(I+1)+((THETA(I)-1.0)
     +         **2)*ETHETA1(I))*(THETA(I+1)-THETA(I))/2
	  END DO
 
C  Initialize parameters for Golden Section Optimization.  For a closed
C  sytem, FLAG=1, For a closed system, FLAG=2

        LEFT = 0.5
	  RIGHT = 80.0 
        NSTEPS = 1000
        TOLG = 1.0D-2
        DFLT = 1.0
        FLAG = 1

        CALL BISECT(LEFT,RIGHT,NSTEPS,TOLG,X,FX)
        PEC = X

	  LEFT=0.5
	  RIGHT=80.0
        FLAG = 2
        CALL BISECT(LEFT,RIGHT,NSTEPS,TOLG,X,FX)
        PEO = X
      
c////////////////////////////////////////////////////////////////////////////////////////
c
c---- Code by Eric J. Oman (1999-Oct-26) begins.
c
      OPEN(UNIT=9, FILE="outpt4.txt", STATUS="unknown")
      OUTPT4_COUNT = NDATA
C      OUTPT4_COUNT = 100
C      OUTPT4_LOW = THETA(1)
C      OUTPT4_HIGH = THETA(NDATA)
C      OUTPT4_THETA_STEP = (OUTPT4_HIGH-OUTPT4_LOW)/DBLE(OUTPT4_COUNT-1)
C      DO I=1,OUTPT4_COUNT
C        OUTPT4_THETA(I) = OUTPT4_LOW + OUTPT4_THETA_STEP*DBLE(I-1)
CC        OUTPT4_E(I) = ETIS(NTANK,OUTPT4_THETA(I))
C      ENDDO      
      DO I=1,OUTPT4_COUNT
        OUTPT4_THETA(I) = THETA(I)
        OUTPT4_E(I) = ETHETA2(I)
      ENDDO      
      WRITE(9,*) 'FILE OUTPT4.TXT'
      WRITE(9,*) 'THE FIRST LINE IS THE NUMBER OF ROWS'
      WRITE(9,*) 'EACH SUBSEQUENT ROW INCLUDES THETA AND E(THETA)'
      WRITE(9,*) '    THETA = DIMENSIONLESS TIME'
      WRITE(9,*) '    E(THETA) = PREDICTED CLOSED MODEL DIMENSIONLESS ',
     & 'CONCENTRATION'
      WRITE(9,*) '-----------------------------------------------------'
      WRITE(9,*) OUTPT4_COUNT
      DO I=1,OUTPT4_COUNT
        WRITE (9,*) OUTPT4_THETA(I), OUTPT4_E(I)
      ENDDO
      CLOSE(9)      
c
c---- Code by Eric J. Oman (1999-Oct-26) ends.
c
c////////////////////////////////////////////////////////////////////////////////////////
c////////////////////////////////////////////////////////////////////////////////////////
c
c---- Code by Eric J. Oman (1999-Oct-26) begins.
c
      OPEN(UNIT=9, FILE="outpt5.txt", STATUS="unknown")
      OUTPT5_COUNT = 100
      OUTPT5_LOW = THETA(1)
      OUTPT5_HIGH = THETA(NDATA)
      OUTPT5_THETA_STEP = (OUTPT5_HIGH-OUTPT5_LOW)/DBLE(OUTPT5_COUNT-1)
      DO I=1,OUTPT5_COUNT
        OUTPT5_THETA(I) = OUTPT5_LOW + OUTPT5_THETA_STEP*DBLE(I-1)
C        OUTPT5_E(I) = ETIS(NTANK,OUTPT5_THETA(I))
        Q1 = 1.0/(2.0*DSQRT(PI*OUTPT5_THETA(I)/PEO))
        Q2 = (1.0-OUTPT5_THETA(I))**2.0*PEO/(4.0*OUTPT5_THETA(I))
        OUTPT5_E(I) = Q1*DEXP(-Q2)
      ENDDO      
      WRITE(9,*) 'FILE OUTPT5.TXT'
      WRITE(9,*) 'THE FIRST LINE IS THE NUMBER OF ROWS'
      WRITE(9,*) 'EACH SUBSEQUENT ROW INCLUDES THETA AND E(THETA)'
      WRITE(9,*) '    THETA = DIMENSIONLESS TIME'
      WRITE(9,*) '    E(THETA) = PREDICTED OPEN MODEL DIMENSIONLESS ',
     & 'CONCENTRATION'
      WRITE(9,*) '-----------------------------------------------------'
      WRITE(9,*) OUTPT5_COUNT
      DO I=1,OUTPT5_COUNT
        WRITE (9,*) OUTPT5_THETA(I), OUTPT5_E(I)
      ENDDO
      CLOSE(9)      
c
c---- Code by Eric J. Oman (1999-Oct-26) ends.
c
c////////////////////////////////////////////////////////////////////////////////////////

C  Print results of above calculations
      
        WRITE(8,1001) VAR,CO
        WRITE(8,1008) TBAR
        WRITE(8,1002) PEC
        WRITE(8,1003) PEO

        WRITE(8,1004)
        DO I=1, NDATA
          WRITE(8,1005) THETA(I), ETHETA1(I), ETHETA2(I), ETHETA3(I)
	  END DO

        WRITE(8,1006)
        DO I=1, ITER1
          WRITE(8,1007) PECS(I),OBJC(I)
	  END DO

        WRITE(8,1009)
        DO I=1, ITER2
          WRITE(8,1007) PEOS(I),OBJO(I)
	  END DO

1001    FORMAT(1X,'Variance : ',T50,E12.6,/,1X,'Initial Concentration 
     +:',T50,E12.6)
1002    FORMAT(1X,'Optimum Peclet number for closed system :',T50,E12.6)
1003    FORMAT(1X,'Optimum Peclet number for open system :',T50,E12.6)
1004    FORMAT(//,1x,'Optimum E(theta) results : ',//,T5,'THETA',T23,
     +  'E(theta)',T43,'E(theta)',T63,'E(theta)',/,T20,'(experimental)',
     +  T40,'(closed model)',T60,'(open model)')
1005    FORMAT(T1,E12.6,T20,E12.6,T40,E12.6,T60,E12.6)
1006    FORMAT(/,1X,'Closed System Peclet Number Optimization',//,T5,
     +  'Pe #',T23,'OBJ.FCN')
1007    FORMAT(T1,E12.6,T20,E12.6)
1008    FORMAT(1X,'Mean Residence Time (Tbar) :',T50,E12.6)
1009    FORMAT(//,1x,'Open System Peclet Number Optimization Results',
     +  //,T5,'Pe #',T23,'OBJ.FCN')

        STOP

      END


C    Function OBJFCN is the object function of residue between the exit age 
C    curve and that of a specific peclet number 
     
      DOUBLE PRECISION FUNCTION OBJFCN(AK)
      
	  INTEGER S
        PARAMETER(S=1600,PI=3.14159265359)
        INTEGER NDATA,FLAG,NIT,ITER,ITER1,ITER2

        DOUBLE PRECISION ETHETA1,ETHETA2,ETHETA3,THETA,MURT,ALPHA,Q1,Q2,
     +       X1,X2,RTSAFE,TOL,AK,OBJ,PEC,PEO,PEOS,PECS,OBJO,OBJC
      
        DIMENSION ETHETA1(S),ETHETA2(S),ETHETA3(S),THETA(S),MURT(S),
     +          PEOS(S),PECS(S),OBJO(S),OBJC(S)

        EXTERNAL RTSAFE,FUNCD

	  COMMON/FLAG/FLAG
        COMMON/RTD/NDATA,ITER1,ITER2,ETHETA1,ETHETA2,ETHETA3,THETA,PECS,
     +           PEOS,OBJO,OBJC
        COMMON/NEWT/ALPHA
        COMMON/SECT/ITER

C     Calculate peclet number for close system

        IF(FLAG.EQ.1) THEN

C  Calculate roots of mu at intervals of pi for use in Thomas and McKee     
C  solution using Newton-Raphson/Bisection Failsafe method
       
          PEC = AK
          ALPHA = 0.5*PEC
          NIT =50
          TOL = 1.0D-8
      
          DO J=1, NIT
            X1 = (J-1)*PI+1.0E-4
            X2 = J*PI-1.0E-4  
            MURT(J) = RTSAFE(FUNCD,X1,X2,TOL)
	    END DO

C   Calculate E(theta) of close system using Thomas an McKee model

          THETA(1) = 0.0055
          Q1 = 0.0
          Q2 = 0.0

          DO I=1, NDATA
            ETHETA2(I) = 0.0
            DO J=1, NIT
              Q1=(ALPHA*DSIN(MURT(J))+MURT(J)*DCOS(MURT(J)))/(ALPHA**
     +           2.0+2.0*ALPHA+MURT(J)**2)
              Q2=ALPHA-((ALPHA**2.0+MURT(J)**2.0)*THETA(I))/(2.0*ALPHA)
              ETHETA2(I)=ETHETA2(I)+2.0*MURT(J)*Q1*DEXP(Q2)
              Q1=0.0
              Q2=0.0
	      END DO 
	    END DO
		 
        ELSE

C  Calculate E(theta) for open system

          THETA(1) = 1.0E-10
          Q1 = 0.0
          Q2 = 0.0
          PEO = AK
      
          DO I=1, NDATA
            Q1 = 1/(2.0*DSQRT(PI*THETA(I)/PEO))
            Q2 = (1-THETA(I))**2.0*PEO/(4.0*THETA(I))
            ETHETA3(I) = Q1*DEXP(-Q2)
	    END DO

        END IF

        OBJFCN = 0.0
        OBJ = 0.0

        DO I=1, NDATA
        
	    IF(ETHETA1(I).EQ.0.0) THEN
	      OBJ = OBJ+0.0
	    ELSE
        
	      IF(FLAG.EQ.1) THEN
              OBJ = OBJ+((ETHETA1(I)-ETHETA2(I))/ETHETA1(I))**2
            ELSE
              OBJ = OBJ+((ETHETA1(I)-ETHETA3(I))/ETHETA1(I))**2
            END IF
	  
	    END IF

	  END DO

        OBJFCN = DSQRT((OBJ)/(NDATA-1))

        IF(FLAG.EQ.1) THEN
          OBJC(ITER) = OBJFCN
          PECS(ITER) = AK
          ITER1 = ITER
        ELSE
          OBJO(ITER) = OBJFCN
          PEOS(ITER) = AK
          ITER2 = ITER
        END IF

C        ITER = ITER+1

        RETURN

      END

C  Function FUNCD contains the object function and its derivation of mu
      
      SUBROUTINE FUNCD(X,F,DF)
      
        DOUBLE PRECISION X,F,DF,ALPHA
        COMMON/NEWT/ALPHA

        F = DCOTAN(X)-0.5*(X/ALPHA-ALPHA/X)
        DF = -1/(DSIN(X)**2)-0.5*(1/ALPHA+ALPHA/(X**2))
      
        RETURN

      END


      FUNCTION rtsafe(funcd,x1,x2,xacc)

        INTEGER MAXIT
        DOUBLE PRECISION rtsafe,x1,x2,xacc
        EXTERNAL funcd
        PARAMETER (MAXIT=1000)
        INTEGER j
        DOUBLE PRECISION df,dx,dxold,f,fh,fl,temp,xh,xl

        call funcd(x1,fl,df)
        call funcd(x2,fh,df)

        if((fl.gt.0..and.fh.gt.0.).or.(fl.lt.0..and.fh.lt.0.))pause
     *  'root must be bracketed in rtsafe'
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

        do 11 j=1,MAXIT
          if(((rtsafe-xh)*df-f)*((rtsafe-xl)*df-f).ge.0..or. abs(2.*
     *          f).gt.abs(dxold*df) ) then
            dxold=dx
            dx=0.5*(xh-xl)
            rtsafe=xl+dx
            if(xl.eq.rtsafe)return
          else
            dxold=dx
            dx=f/df
            temp=rtsafe
            rtsafe=rtsafe-dx
            if(temp.eq.rtsafe)return
          endif

          if(abs(dx).lt.xacc) return

          call funcd(rtsafe,f,df)

          if(f.lt.0.) then
            xl=rtsafe
          else
            xh=rtsafe
          endif
11      continue
        pause 'rtsafe exceeding maximum iterations'
        return

      END
C  (C) Copr. 1986-92 Numerical Recipes Software 41.921'L3.


C   Subroutine BISECT calculates the peclect of a system by bi-section 
C   method

	SUBROUTINE BISECT(A,B,MAXIT,TOL,X,FX)

        IMPLICIT DOUBLE PRECISION(A-H,O-Z)
	  DOUBLE PRECISION DF,TEMP1,TEMP2,MID,TEMPFX1,TEMPFX2    
        INTEGER ITER
	  EXTERNAL OBJFCN
	  COMMON/SECT/ITER

	  TEMPFX1 = OBJFCN(A)
	  TEMPFX2 = OBJFCN(B)

	  ITER = 1    
             
	  DELT = 0.001

	  DO WHILE(ITER.LT.MAXIT)

	    MID = (A+B)/2.0
	    TEMP1 = MID + DELT
	    TEMP2 = MID - DELT
	    FX1 = OBJFCN(TEMP1)
	    FX2 = OBJFCN(TEMP2)
	    DF=(FX1**2 - FX2**2)/(2.*DELT)
	    ITER=ITER+1

          IF (DF.GT.0.0) THEN
	      B=MID
	    ELSE IF (DF.LT.0.0) THEN
	      A=MID
	    END IF 

	    IF ((ABS(B-A).LE.TOL).OR.(ABS(DF).LE.1E-5)) EXIT

	  END DO  

	  IF (ABS(DF).GT.1E-2) THEN
	    IF ((FX1.GE.TEMPFX1).AND.(FX2.GE.TEMPFX2)) THEN
	      WRITE(8,802) ITER
            RETURN
	    END IF
	  ELSE 
          X = MID
	    FX = OBJFCN(X)
	  END IF


802     FORMAT(//,1X,'** ERROR:SUBROUTINE BISECTION CANNOT FIND THE ROOT 
     &  AFTER ',I6,'ITERATIONS;',/,12X,'ENLARGE THE RANGE OF PECLET',//)
	
	RETURN

      END

