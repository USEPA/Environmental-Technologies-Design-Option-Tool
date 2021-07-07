REM  PROGRAM:  PULL.BAS

REM  DEVELOPED BY ANDREW W. LOLL
REM  3/96
REM QBASIC VERSION 1.0

REM  THIS PROGRAM PULLS THE DESIRED CHEMICALS OUT OF MY FORTRAN FORMATTED
REM  INPUT FILES.

REM********************************************************************

DIM a$(100)

a = 0
count = 0

CLS

OPEN "C:\FIT\FITT95A.DAT" FOR INPUT AS #1
OPEN "C:\FIT\STUFF.OUT" FOR OUTPUT AS #2

25 INPUT "DO YOU WANT THE EXPERIMENTAL DATA(enter 'y' or 'n')?  ", c$
REM   IF LTRIM$(c$) <> "y" OR LTRIM$(c$) <> "n" THEN
REM      GOTO 25
REM   END IF
50 a = a + 1
   INPUT "ENTER CAS# OF CHEMICAL DESIRED(enter 'n' if done): ", a$
   IF a$ = "n" THEN
      a = a - 1
      GOTO 75
   END IF
   cas$(a) = a$
   GOTO 50

75 amax = a


iter:

IF EOF(1) THEN

    CLOSE (1)
    CLOSE (2)
    GOTO done:

END IF

INPUT #1, out1$, out2$, out3$

b = 0

DO UNTIL b = amax
   b = b + 1
   IF VAL(out1$) = VAL(cas$(b)) THEN
      cflag = 1
      PRINT , "Working on CAS#:  " + cas$(b)
      PRINT #2, out1$ + "," + out2$ + "," + out3$
      dmax = VAL(out3$)
      FOR i = 1 TO dmax
         IF c$ = "y" THEN
            LINE INPUT #1, data$
            PRINT #2, data$
         ELSE
            LINE INPUT #1, data$
         END IF
      NEXT
      GOTO iter:
   ELSE
      cflag = 0
   END IF
LOOP
IF cflag = 0 THEN
   smax = VAL(out3$)
   FOR j = 1 TO smax
      LINE INPUT #1, data$
   NEXT
END IF
GOTO iter:

done:
END

