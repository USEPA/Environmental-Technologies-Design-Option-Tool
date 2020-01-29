
'0123456789012345678901234567890       (STRINGBASE0 NOTATION)
'0000000000111111111122222222223
'CA****-*****-*****-*****-*****
'..XX.X-XXXXX-XXXXX-.XXXX-X.XXX
'123456789012345678901234567890        (STRINGBASE1 NOTATION)
'000000000111111111122222222223


'SHOULD HAVE:
'2,3,5,7,8,9,10,11,13,14,15,16,17,20,21,22,23,25,27,28,29
'CHECK:
'2,3,5,7,8,9,10,11,13,14,15,16,17,20,21,22,23,25,27,28,29

DIM x(21)
x(1) = 3
x(2) = 4
x(3) = 6
x(4) = 8
x(5) = 9
x(6) = 10
x(7) = 11
x(8) = 12
x(9) = 14
x(10) = 15
x(11) = 16
x(12) = 17
x(13) = 18
x(14) = 21
x(15) = 22
x(16) = 23
x(17) = 24
x(18) = 26
x(19) = 28
x(20) = 29
x(21) = 30

DIM y(21)

j = 0
FOR i = 1 TO 10000
  t = INT(RND(1) * 21) + 1
  IF (x(t) > 0) THEN
    j = j + 1
    y(j) = x(t)
    x(t) = 0
  END IF
  IF (j = 21) THEN EXIT FOR
NEXT i

OPEN "scatter.out" FOR OUTPUT AS #1
FOR i = 1 TO 21
  'NOTE: CONVERT STRINGBASE1 NOTATION TO STRINGBASE0 NOTATION.
  PRINT #1, y(i) - 1
NEXT i
CLOSE #1

