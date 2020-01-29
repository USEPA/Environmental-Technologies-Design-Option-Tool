REM PROGRAM: COUNT2.BAS

REM DIPPR801 Dataform Program
REM Designed by Andrew W. Loll
REM QBASIC Version 1.00

REM This program counts the number of compounds in my FORTRAN formatted
REM input file.

REM**************************************************************************

CLS

count = 0
comp = 0
vpcount = 0
badchem = 0
goodchem = 0
chem = 1

OPEN "c:\FIT\fitt95-2.dat" FOR INPUT AS #1
OPEN "c:\FIT\FITT2.DAt" FOR OUTPUT AS #2

main:

GOSUB screenp:

iter:

IF EOF(1) THEN

     GOTO tempsort:
       
END IF

INPUT #1, out$, OUT2$, out3$

REM  PRINT , OUT2$
IF count < 1 THEN
        
         dippr = VAL(out$)
        lastdippr = VAL(out$)
        PRINT #2, out$
       
END IF

count = count + 1

IF lastdippr <> VAL(out$) THEN

        vpcount = 0
        lastdippr = VAL(out$)
        dippr = VAL(out$)

END IF

IF VAL(out$) <> dippr THEN

        vpcount = 0
    
END IF







IF comp <> dippr THEN

        comp = dippr
        chem = chem + 1
        IF VAL(out3$) = -1 THEN

             badchem = badchem + 1
             PRINT #3, badchem; SPC(5); out$
             GOTO finish:

        END IF

        goodchem = goodchem + 1
        GOSUB screenp:

finish:

END IF

GOTO iter:

tempsort:

IF EOF(1) THEN
    
     chem = chem - 1
     PRINT , chem
     CLOSE 1
     GOTO andy:

END IF

GOTO iter:

andy:


END

screenp:

CLS
PRINT " Working on ext1.dat"
PRINT chem
PRINT #2, goodchem; SPC(5); out$
RETURN

novp:

IF MID$(VP$(z - 1), 1, 11) = "           " THEN z = z - 1

RETURN


     



