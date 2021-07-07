      PROGRAM LINTEST

C  THIS PROBLEM FEEDS IN THE NECESSARY DATA TO FIT.FOR FOR MY LINEAR MODEL 
C  TEST CASE AS DESCRIBED IN EX9-19.XLS.

      IMPLICIT INTEGER (A-G,M-P)
      REAL TP(830,900)
      DIMENSION NPOINT(830,2), FRAGLIN(830,108)
      EXTERNAL LINPREP

      OPEN (21, FILE='FRAGTEST.PRN',STATUS='UNKNOWN',
     &          ACCESS='SEQUENTIAL',FORM='FORMATTED') 


      MAXFRG = 3
      CMAX = 17
      TP(1,1) = 298.15
      TP(1,2) = 23.90678559
      TP(2,1) = 285.15
      TP(2,2) = 27.43549545
      TP(2,3) = 305.15
      TP(2,4) = 27.72137661
      TP(2,5) = 355.15
      TP(2,6) = 28.30423397
      TP(3,1) = 340.15
      TP(3,2) = 32.58617083
      TP(4,1) = 345.15
      TP(4,2) = 37.81391274
      TP(4,3) = 352.15
      TP(4,4) = 37.93235132
      TP(5,1) = 368.15
      TP(5,2) = 44.25688113
      TP(6,1) = 377.15
      TP(6,2) = 51.50258377
      TP(7,1) = 390.15
      TP(7,2) = 60.06646802
      TP(8,1) = 402.15
      TP(8,2) = 70.05086937
      TP(9,1) = 410.15
      TP(9,2) = 81.5757769
      TP(10,1) = 426.15
      TP(10,2) = 95.37477686
      TP(11,1) = 285.15
      TP(11,2) = 7.710079461
      TP(12,1) = 292.15
      TP(12,2) = 8.935901289
      TP(13,1) = 315.15
      TP(13,2) = 10.44812293
      TP(14,1) = 326.15
      TP(14,2) = 12.15205248
      TP(14,3) = 355.15
      TP(14,4) = 12.31437503
      TP(15,1) = 345.15
      TP(15,2) = 14.19821193
      TP(16,1) = 358.15
      TP(16,2) = 16.54906741
      TP(17,1) = 402.15
      TP(17,2) = 19.56794613

      NPOINT(1,1) = 590181
      NPOINT(1,2) = 1
      NPOINT(2,1) = 627203
      NPOINT(2,2) = 3
      NPOINT(3,1) = 7688213 
      NPOINT(3,2) = 1
      NPOINT(4,1) = 6443921
      NPOINT(4,2) = 2
      NPOINT(5,1) = 7642151
      NPOINT(5,2) = 1
      NPOINT(6,1) = 124118
      NPOINT(6,2) = 1
      NPOINT(7,1) = 872059
      NPOINT(7,2) = 1
      NPOINT(8,1) = 821954
      NPOINT(8,2) = 1
      NPOINT(9,1) = 112414
      NPOINT(9,2) = 1
      NPOINT(10,1) = 2437561
      NPOINT(10,2) = 1
      NPOINT(11,1) = 106978
      NPOINT(11,2) = 1
      NPOINT(12,1) = 109660
      NPOINT(12,2) = 1
      NPOINT(13,1) = 110543
      NPOINT(13,2) = 1
      NPOINT(14,1) = 142825
      NPOINT(14,2) = 2
      NPOINT(15,1) = 111659
      NPOINT(15,2) = 1
      NPOINT(16,1) = 111842
      NPOINT(16,2) = 1
      NPOINT(17,1) = 124185
      NPOINT(17,2) = 1
      
      DO 6000, I=1,17
         READ (21,*) CAS, (FRAGLIN(I,J),J=1,3)
6000  CONTINUE
      CALL LINPREP(FRAGLIN, TP, NPOINT, MAXFRG, CMAX)
      CLOSE (21)
      RETURN
      END
