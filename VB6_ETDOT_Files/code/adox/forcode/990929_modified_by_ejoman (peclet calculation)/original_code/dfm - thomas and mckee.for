$DEBUG
      PROGRAM DFMRTD
      
      PARAMETER(S=1600,PI=3.14159265359)
      INTEGER NDATA,NDATA,FLAG,NIT
      REAL*8 CO,TBAR,ETHETA1,ETHETA2,ETHETA3,T,C,CT,THETA,VAR,
     +       INTV,XTHETA1,XTHETA2,CNOT,PEC,PEO,MURT,ALPHA,Q1,Q2,
     +       X1,X2,RTSAFE,TOL
      
      DIMENSION ETHETA1(S),ETHETA2(S),ETHETA3(S),T(S),C(S),XTHETA1(S),
     +XTHETA2(S),THETA(S),MURT(S)
      CHARACTER *20 INDATA, OUTDATA    
      EXTERNAL RTSAFE,FUNCD
      COMMON/NEWT/VAR,ALPHA
      

C  This program takes the data from a pulse dye study andcalculates the 
C  exit age distribution.  It then uses this data to calculate the varience, 
C  from C  which the Peclet numbers for an open and closed system are 
C  determined. Using these numbers with the appropriate model, E(theta) 
C  curves are predicted for both open and closed systems
C
C  PROGRAMMED BY :
C                    Louis Kindt
C                    Department of Chemical Engineering
C                    Michigan Technological University
C
C  Subroutine RTSAFE used in determining necessary roots using a Newton-
C  Raphson/Bisection Failsafe method was taken from  "Numerical Recipes"
C  Cambridge University Press,1994.
C
C  NOMENCLATURE
C    INDATA - input file of the form NDATA/T(i), C(i)
C    OUTDATA - output file for results of all calculations
C    NDATA - number of data points of experimental data
C    T(i) - experimental time data (maximum length = 1600)
C    C(i) - experimental concentration data (maximum length = 1600)
C    CO - initial concentration
C    TBAR - mean residence time
C    THETA - dimensionless time
C    VAR - varience of exit age distribution
C    X(1) - peclet number for a closed system
C    X(2) - peclet number for an open system
C    X(3) - mu for use in Thomas and McKee solution
C    ETHETA1(i) = E(theta) for experimental data
C    ETHETA2(i) = predicted E(theta) for closed system
C    ETHETA3(i) = predicted E(theta) for open system
C    XTHETA1(i) = desired theta time steps for closed system
C    XTHETA2(i) = desired theta time steps for open system

      OPEN(UNIT=7, FILE="dfminput.txt", STATUS="unknown")
      OPEN(UNIT=8, FILE="dfmoutpt.txt", STATUS="unknown")
 	 
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

      DO 30 I=1, NDATA
         ETHETA1(I) = C(I)/CO
         THETA(I) = T(I)/TBAR
30    CONTINUE

C  Calculation of varience (sigma**2) of the exit age distribution

      VAR = 0.0

      DO 40 I=1, NDATA-1
         VAR = VAR+(((THETA(I+1)-1.0)**2)*ETHETA1(I+1)+((THETA(I)-1.0)
     +         **2)*ETHETA1(I))*(THETA(I+1)-THETA(I))/2
40    CONTINUE
 
C  Solve for the Peclet number using the above varience for a closed
C  system and an open system using Newton-Raphson/Bisection Failsafe 
C  technique for root finding.  
C  For closed system : FLAG=1, open system : FLAG=2, mu : FLAG = 3

C  Set interval for root search and tolerance for root finding method
 
      X1 = 1.0E-8
      X2 = 1.0E10
      TOL = 1.0E-8

      FLAG = 1
      PEC = RTSAFE(FUNCD,X1,X2,TOL,FLAG)

      FLAG = 2
      PEO = RTSAFE(FUNCD,X1,X2,TOL,FLAG)

C  Calculate roots of mu at intervals of pi for use in Thomas and McKee     
C  solution using Newton-Raphson/Bisection Failsafe method
       
      ALPHA = 0.5*PEC
      FLAG = 3
      NIT =50

      DO 50 J=1, NIT
         X1 = (J-1)*PI+1.0E-4
         X2 = J*PI-1.0E-4  
         MURT(J) = RTSAFE(FUNCD,X1,X2,TOL,FLAG)
50    CONTINUE

C  Calculate E(theta) using Thomas an McKee model

      INTV = THETA(NDATA)/NDATA 
      XTHETA1(1) = 0.0055
      Q1 = 0.0
      Q2 = 0.0

      DO 60 I=1, NDATA
         ETHETA2(I) = 0.0
         DO 65 J=1, NIT
            Q1=(ALPHA*DSIN(MURT(J))+MURT(J)*DCOS(MURT(J)))/(ALPHA**2.0+
     +         2.0*ALPHA+MURT(J)**2)
            Q2=ALPHA-((ALPHA**2.0+MURT(J)**2.0)*XTHETA1(I))/(2.0*ALPHA)
            ETHETA2(I)=ETHETA2(I)+2.0*MURT(J)*Q1*DEXP(Q2)
            Q1=0.0
            Q2=0.0
65       CONTINUE
         XTHETA1(I+1)=XTHETA1(I)+INTV
60    CONTINUE

C  Calculate E(theta) for open system

      XTHETA2(1) = 1.0E-10
      Q1 = 0.0
      Q2 = 0.0
      
      DO 70 I=1, NDATA
         Q1 = 1/(2.0*DSQRT(PI*XTHETA2(I)/PEO))
         Q2 = (1-XTHETA2(I))**2.0*PEO/(4.0*XTHETA2(I))
         ETHETA3(I) = Q1*DEXP(-Q2)
         XTHETA2(I+1) = XTHETA2(I)+INTV
70    CONTINUE

C  Print results of above calculations
      WRITE(8,1000) INDATA
      WRITE(8,1001) VAR,CO
      WRITE(8,1008) TBAR
      WRITE(8,1002) PEC
      WRITE(8,1003) PEO

      WRITE(8,1004)
      DO 80, I=1, NDATA
         WRITE(8,1005) THETA(I), ETHETA1(I)
80    CONTINUE

      WRITE(8,1006)
      DO 90, I=1, NDATA
         WRITE(8,1005) XTHETA1(I), ETHETA2(I)
90    CONTINUE

      WRITE(8,1007)
      DO 100, I=1, NDATA
         WRITE(8,1005) XTHETA2(I), ETHETA3(I)
100   CONTINUE

1000  FORMAT(1X,'DISPERSED FLOW MODEL RESULTS FOR OPEN AND CLOSED SYSTEM
     +S',//,1X,'Input File :',T42,A20)
1001  FORMAT(1X,'Varience :',T40,E12.6,/,1X,'Initial Concentration :'
     +,T40,E12.6)
1002  FORMAT(1X,'Peclet number for closed system :',T40,E12.6)
1003  FORMAT(1X,'Peclet number for open system :',T40,E12.6)
1004  FORMAT(//,1x,'Experimental E(theta) results : ',//,T9,'THETA',T29,
     +'E(theta)')
1005  FORMAT(T5,E12.6,T25,E12.6)
1006  FORMAT(//,1X,'Model predictions of E(theta) for closed system :'
     +,//,T9,'THETA',T29,'E(theta)')
1007  FORMAT(//,1X,'Model predictions of E(theta) for open system :
     +',//,T9,'THETA',T29,'E(theta)')
1008  FORMAT(1X,'Mean Residence Time (Tbar) :',T40,E12.6)

      STOP
      END

      
      SUBROUTINE FUNCD(X,F,DF,FLAG)
      
      INTEGER FLAG
      REAL*8 X,VAR,F,DF,ALPHA
      COMMON/NEWT/VAR,ALPHA

      IF(FLAG.EQ.1) THEN
        F = VAR-(2/X)+(2.0/(X**2.0))*(1-DEXP(-X))
        DF = 2/X**2+2*(DEXP(-X))/X**2-4*(1-DEXP(-X))/X**3
      ELSE IF(FLAG.EQ.2) THEN
        F = VAR-(2/X)-(8.0/(X**2.0))
        DF = 2/X**2+16/X**3
      ELSE
        F = DCOTAN(X)-0.5*(X/ALPHA-ALPHA/X)
        DF = -1/(DSIN(X)**2)-0.5*(1/ALPHA-ALPHA/(X**2))
      END IF
              
      RETURN
      END


      FUNCTION rtsafe(funcd,x1,x2,xacc,p)
      INTEGER MAXIT
      REAL*8 rtsafe,x1,x2,xacc
      EXTERNAL funcd
      PARAMETER (MAXIT=1000)
      INTEGER j,p
      REAL*8 df,dx,dxold,f,fh,fl,temp,xh,xl
      call funcd(x1,fl,df,p)
      call funcd(x2,fh,df,p)
      if((fl.gt.0..and.fh.gt.0.).or.(fl.lt.0..and.fh.lt.0.))pause
     *'root must be bracketed in rtsafe'
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
      call funcd(rtsafe,f,df,p)
      do 11 j=1,MAXIT
        if(((rtsafe-xh)*df-f)*((rtsafe-xl)*df-f).ge.0..or. abs(2.*
     *f).gt.abs(dxold*df) ) then
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
        call funcd(rtsafe,f,df,p)
        if(f.lt.0.) then
          xl=rtsafe
        else
          xh=rtsafe
        endif
11    continue
      pause 'rtsafe exceeding maximum iterations'
      return
      END
C  (C) Copr. 1986-92 Numerical Recipes Software 41.921'L3.

