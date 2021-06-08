REM PROGRAM: FORM2.BAS

REM DIPPR801 Dataform Program
REM Designed by Andrew W. Loll
REM QBASIC Version 1.00

REM This routine changes the data files from the DIPPR801 format to give a
REM list of the extracted compound

CLS

OPEN "c:\fit\ext95.dat" FOR INPUT AS #1
OPEN "c:\fit\vplist95.dat" FOR OUTPUT AS #2

count = 0
comp = 0

PRINT #2, "  CAS#   "; SPC(6); "Common Name"
PRINT #2, "---------"; SPC(6); "-----------"
PRINT #2,

iter:

IF EOF(1) THEN

        CLOSE 1
        GOTO andy:
       
END IF

LINE INPUT #1, out$

IF count < 1 THEN

        dippr = VAL(MID$(out$, 1, 4))
        lastdippr = VAL(MID$(out$, 1, 4))
       
END IF

count = count + 1

IF lastdippr <> VAL(MID$(out$, 1, 4)) THEN

        lastdippr = VAL(MID$(out$, 1, 4))
        dippr = VAL(MID$(out$, 1, 4))

END IF

IF VAL(MID$(out$, 1, 4)) <> dippr GOTO iter:



temp$ = MID$(out$, 5, 4)

REM


IF temp$ = "CNAM" THEN

        CNAME$ = MID$(out$, 11, 60)
                          
END IF

IF temp$ = "CASN" THEN

        CAS$ = MID$(out$, 11, 9)
        PRINT #2, RTRIM$(CAS$); SPC(6); LTRIM$(CNAME$)

        comp = comp + 1
        PRINT comp

END IF

GOTO iter:

andy:

PRINT #2, "The number of compounds extracted is: ", comp

CLOSE 2

END






