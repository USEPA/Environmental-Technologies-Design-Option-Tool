      PROGRAM EXAMP3

      REAL X(3),XGUESS(3),S,FTOL,FVALUE,Y(5),IND(5)
      INTEGER MAXFCN,MPAR,NDAT
      EXTERNAL FCN, UMPOL
      COMMON /A/ Y,NDAT,IND

      DATA XGUESS(1),XGUESS(2),XGUESS(3)/4.0,3.0,1.0/
      DATA Y(1),Y(2),Y(3),Y(4),Y(5)/0.435,5.224,8.9934,22.3452,100.342/
      DATA IND(1),IND(2),IND(3),IND(4),IND(5)/3.0,5.0,7.0,9.0,11.0/
      S = 5.0
      MAXFCN = 10000
      MPAR = 3
      NDAT = 5
      FTOL = 1.0E-7

      DO 10, I=1,3
         PRINT *, I, ' XGUESS: ',XGUESS(I)
10    CONTINUE

      CALL UMPOL(FCN,MPAR,XGUESS,S,FTOL,MAXFCN,X,FVALUE)

      PRINT *, 'THE FINAL VALUES ARE: '
      PRINT *
      DO 20, J=1,3
         PRINT *, J,' X: ',X(J)
20    CONTINUE
      PRINT *
      PRINT *,'THE FINAL OBJECTIVE FUNCTION IS: ',FVALUE

      END

****************************************************************

      SUBROUTINE FCN(NN,XX,FF)
      
      REAL XX(3),FF,Y(5),IND(5),VALUE
      INTEGER NN,NDAT
      COMMON /A/ Y,NDAT,IND
      
      DO 25,JJ=1,NN
         PRINT *,JJ,' XX: ',XX(JJ)
25    CONTINUE
      FF=0.0
      DO 30, K = 1,NDAT

         VALUE = XX(1)*((SQRT(IND(K)+XX(2))**2.)/(1.0+XX(3))) 
         FF=FF+(Y(5)-VALUE)**2

30    CONTINUE
      FF=SQRT(FF/NDAT)

      RETURN
      END
