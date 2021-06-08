      PROGRAM ADDCOEFF

C  THIS PROGRAM FINDS MATCHES THE CHEMICALS UP FROM SINGLE.DAT TO THE 
C  APPROPRIATE CAS NUMBER IN UNIFACT.DAT AND ADDS UP THE MODEL COEFFICIENTS
C  FOR THOSE CHEMICALS FROM THE SPECIFIED SEED FILE.

      INTEGER I,IMAX,CHEM(820),FRAG(1430,22),CAS(1430),UNIMAX,IA,IB,
     &    IC,ID,JA
      DOUBLE PRECISION SEED(115,2),ASUM,BSUM 
      CHARACTER*20 DFORM
      CHARACTER*12 DCAS
      CHARACTER*4  DDUM(22)

      CFIT = 0

      OPEN(21,FILE='single.dat',STATUS='UNKNOWN',ACCESS='SEQUENTIAL',
     &        FORM='FORMATTED')
      OPEN(22,FILE='unifact.dat',STATUS='UNKNOWN',ACCESS='SEQUENTIAL',
     &        FORM='FORMATTED')
      OPEN(23,FILE='seed.dat',STATUS='UNKNOWN',ACCESS='SEQUENTIAL',
     &        FORM='FORMATTED')
      OPEN(24,FILE='coeff.out',STATUS='UNKNOWN',ACCESS='SEQUENTIAL',
     &        FORM='FORMATTED')
      OPEN(25,FILE='badchem.out',STATUS='UNKNOWN',ACCESS='SEQUENTIAL',
     &        FORM='FORMATTED')
     
      I=0

10    I=I+1
      READ(21,*,ERR=19) CHEM(I)
      GOTO 10
19    IMAX=I-1
      PRINT *,'IMAX: ',IMAX
      DO 15,J=1,115
         READ(23,*) SEED(J,1),SEED(J,2)
15    CONTINUE
   
      READ(22,*) DCAS,DFORM,(DDUM(K),K=1,22)
      IA=0
20    IA=IA+1
      READ(22,*,ERR=29) CAS(IA),DFORM,(FRAG(IA,K),K=1,22)
      GOTO 20
29    UNIMAX=IA-1
      CLOSE(21)
      CLOSE(22)
      CLOSE(23)

      DO 30, IB=1,IMAX
         DO 40, IC=1,UNIMAX
            IF (CAS(IC).EQ.CHEM(IB)) THEN
               FMAX=FRAG(IC,2)
               JA = 0
               ASUM = 0.0
               BSUM = 0.0
               DO 50, ID=1,FMAX
                  JA=ID*2+1
                  IF(SEED(FRAG(IC,JA),1).EQ.0.) GOTO 35
                  ASUM=ASUM+DFLOAT(FRAG(IC,JA+1))*SEED(FRAG(IC,JA),1)
                  BSUM=BSUM+DFLOAT(FRAG(IC,JA+1))*
     &                 SEED(FRAG(IC,JA),2)
C     &                 SEED(FRAG(IC,JA),2)*100.
50             CONTINUE
               WRITE(24,*) CHEM(IB),',',ASUM,',',BSUM
               GOTO 30
            END IF
40       CONTINUE
         WRITE(25,*) 'CAS #: ',CHEM(IB),'IS NOT IN UNIFACT.DAT'
35       CONTINUE
         WRITE(25,*) 'CAS #: ',CHEM(IB),'HAS AN INVALID FRAGMENT'
30    CONTINUE
      CLOSE(24)
      CLOSE(25)
      END
