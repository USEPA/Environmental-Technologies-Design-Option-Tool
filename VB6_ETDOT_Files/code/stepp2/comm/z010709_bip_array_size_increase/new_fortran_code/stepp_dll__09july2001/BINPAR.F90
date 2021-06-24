!CC  ****************************************************************************
!CC  *                                                                          *
!CC  *               LOADS UNIFAC BINARY INTERACTION PARAMETERS                 *
!CC  *                                                                          *
!CC  ****************************************************************************

      SUBROUTINE BINPAR (MDL,MGSG,AI,RI,QI,FMW,FVB)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::BINPAR
!MS$ ATTRIBUTES ALIAS:'_BINPAR@28':: BINPAR
!MS$ ATTRIBUTES REFERENCE::MDL,MGSG,AI,RI,QI,FMW,FVB

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
!CC-----Modified David R. Hokanson 7/9/01 for STEPP2
!CC-----   Increased dimensioning for new binary interaction parameter databases      
!CC      PARAMETER (LA=32,MA=53,NA=96)
      PARAMETER(LA=32,MA=58,NA=116)
!CC--------End Modified David R. Hokanson 7/9/01 for STEPP2
      

      DIMENSION AI(MA,MA),RI(NA),QI(NA),MGSG(NA),IJTR(NA)
      DIMENSION FMW(NA),FML(NA),FVB(NA),BB(LA)

!CC      OPEN (UNIT = 9, FILE = 'C:\STEPP\RANDQ.DAT')
      OPEN (UNIT = 9, FILE = 'RANDQ.DAT')

      DO 10 J=1,NA

           READ(9,*) MGSG(J),IJTR(J),RI(J),QI(J),FMW(J),FML(J),FVB(J)

  10  CONTINUE

      IF (MDL-2) 20,30,70

!CC  20  OPEN (UNIT = 10, FILE = 'C:\STEPP\AVLE.DAT')
  20  OPEN (UNIT = 10, FILE = 'AVLE.DAT')

      GOTO 80

!CC  30  OPEN (UNIT = 10, FILE = 'C:\STEPP\ALLE.DAT')
  30  OPEN (UNIT = 10, FILE = 'ALLE.DAT')
  
      DO 40 J=1,MA
 
           DO 40 K=1,MA

                AI(J,K) = 99999.0D0

                IF (K.NE.J) GOTO 40

                AI(J,J) = 0.0D0

  40  CONTINUE

      DO 50 J=1,LA

            READ(10,*) (BB(L),L=1,LA)

            DO 50 K=1,LA
  
                 IF (K.EQ.J) GOTO 50

                 AI(IJTR(J),IJTR(K))=BB(K)

  50  CONTINUE

      GOTO 100

!CC  70  OPEN (UNIT = 10, FILE = 'C:\STEPP\AENV.DAT')
  70  OPEN (UNIT = 10, FILE = 'AENV.DAT')

  80  DO 90 I=1,MA

           READ(10,*) (AI(I,J),J=1,MA)

  90  CONTINUE

 100  CLOSE(9)
      CLOSE(10)

      END


