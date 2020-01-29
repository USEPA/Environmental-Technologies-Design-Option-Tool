CLS

OPEN "c:\fit\stat2.808" FOR INPUT AS #1
OPEN "c:\fit\fitset.out" FOR OUTPUT AS #3

iter:
IF EOF(1) THEN
   CLOSE (1)
   CLOSE (3)
   GOTO done:
END IF

INPUT #1, CAS1$
OPEN "c:\fit\reserve.dat" FOR INPUT AS #2
check:
IF EOF(2) THEN
   PRINT #3, CAS1$
   CLOSE (2)
   GOTO iter:
END IF
INPUT #2, CAS2$
IF CAS1$ = CAS2$ THEN
   CLOSE (2)
   GOTO iter:
ELSE
   GOTO check:
END IF

done:

STOP
END

