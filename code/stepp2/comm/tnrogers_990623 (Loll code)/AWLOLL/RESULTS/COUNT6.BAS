REM PROGRAM: COUNT6.BAS

REM DIPPR801 Dataform Program
REM Designed by Andrew W. Loll
REM QBASIC Version 1.00

REM THIS PROGRAM PRINTS ALL OF THE FIRST LINES OF THE CHEMICALS IN MY FORTRAN
REM FORMATTED INPUT FILES SO THAT I CAN SORT THEM IN A DATABASE OR
REM SPREADSHEET.


REM**************************************************************************

CLS

count = 0
comp = 0
vpcount = 0
badchem = 0
goodchem = 0
chem = 1

OPEN "c:\fit\fitt95a.dat" FOR INPUT AS #1
OPEN "c:\fit\fitt95a.out" FOR OUTPUT AS #2

main:

GOSUB screenp:

iter:

count = 0

IF EOF(1) THEN

     GOTO tempsort:
       
END IF

DO

INPUT #1, out$, out2$, out3$

cut$ = LEFT$(out2$, 3)
cut2$ = RIGHT$(cut$, 2)

IF count < 1 THEN
     
     IF VAL(out3$) = -1 OR VAL(out3$) = 0 THEN

        GOTO iter:

     END IF

     PRINT #2, out$ + "," + cut2$ + "," + out3$
     datpoint = VAL(out3$)
     count = count + 1
     GOTO schwen:

END IF

count = count + 1

REM PRINT #2, out$ + "," + out2$ + "," + out3$
    
schwen:

LOOP UNTIL count > datpoint

GOTO iter:

tempsort:

IF EOF(1) THEN
    
     chem = chem - 1
     PRINT , chem
     CLOSE 1
     GOTO andy:

END IF

andy:

CLOSE 2

END

screenp:

CLS
PRINT " Working on ext1.dat"
PRINT chem
REM PRINT #2, goodchem; SPC(5); out$
RETURN

