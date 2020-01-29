!CC****************************************************************************
!CC
!CC                                  FGRP
!CC                          ORDER FUNCTIONAL GROUPS
!CC
!CC Output Variables:
!CC    JERR =     Error Flag
!CC
!CC Input Variables:
!CC    NC =
!CC    NG =
!CC
!CC Authors:  M. Miller, T. Rogers, D. Hokanson (4/4/94)
!CC
!CC****************************************************************************

      SUBROUTINE FGRP (NC,NG,JERR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::FGRP
!MS$ ATTRIBUTES ALIAS:'_FGRP@12':: FGRP
!MS$ ATTRIBUTES REFERENCE::NC,NG,JERR

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      PARAMETER  (MA=53,NA=96,ND=10)

      DIMENSION  NGM(10),NY(10,20),JH(NA),IH(20)

      COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
      COMMON /GROUP/ MS(10,10,2),NMAX
      COMMON /UNI/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),R(10),P(10,10)
      COMMON /LIMITS/ TOL,IMAX
      COMMON /ERR/ ERRMAT(30),ERRNUM

!CC    -- INITIALIZE VARIABLES

      IC = 1
      NK = NC

      DO 10 I=1,10

           DO 10 J=1,NK

                QT(I,J) = 0.0D0
                RT(I,J) = 0.0D0

  10  CONTINUE

      DO 90 I=1,NA

           JH(I) = 0

  90  CONTINUE

      DO 160 I=1,NK

           DO 150 J=1,NMAX

                IF (MS(I,J,1).EQ.0) GOTO 160

                IH(IC) = MS(I,J,1)

                IF (IC.EQ.1) GOTO 140

                IF (IH(IC).EQ.IH(IC-1)) GOTO 150

                IF (IH(IC).GT.IH(IC-1)) GOTO 140

                IF (IC.GT.2) GOTO 110

                IHH = IH(1)
                IH(1) = IH(2)
                IH(2) = IHH

                GOTO 140

 110            I1 = IC-1

                DO 130 I2=1,I1

                     IF (IH(IC).GT.IH(I2)) GOTO 130

                     IF (IH(IC).EQ.IH(I2)) GOTO 150

                     I4 = IC-I2

                     DO 120 I3=1,I4

                          IH(IC+1-I3) = IH(IC-I3)

 120                 CONTINUE

                     IH(I2) = MS(I,J,1)
 
 130            CONTINUE

 140            IC = IC+1

                IF (IC.GT.20) THEN
 
                     JERR = -1
                     CALL ERROR (ERRMAT,ERRNUM,5)
                     RETURN

                END IF 

 150       CONTINUE

 160  CONTINUE

      IC = IC-1

      DO 170 I=1,IC

           JH(IH(I)) = I

 170  CONTINUE

      DO 180 I=1,10

           DO 180 J=1,20

                NY(I,J) = 0

 180  CONTINUE

      DO 200 I=1,NK

           DO 190 J=1,10

                IF (MS(I,J,1).EQ.0) GOTO 200 

                N1 = MS(I,J,1)
                N2 = MS(I,J,2)

                IF (N1.EQ.0) GOTO 200

                N3 = JH(N1)
                NY(I,N3) = N2

 190       CONTINUE

 200  CONTINUE

      I = 0
      NGMGL = 0

      DO 210 K=1,IC

           NSG = IH(K)
           NGMNY = MGSG(NSG)
           IF (NGMNY.NE.NGMGL) I=I+1
           NGM(I) = NGMNY
           NGMGL = NGMNY

      DO 210 J=1,NK

           RT(I,J) = RT(I,J)+DBLE(NY(J,K))*RI(NSG)
           QT(I,J) = QT(I,J)+DBLE(NY(J,K))*QI(NSG)

 210  CONTINUE

      NG = I

      DO 220 I=1,NG

           DO 220 J=1,NG

                NI = NGM(I)
                NJ = NGM(J)
                AVAL = AI(NI,NJ)

                IF (DABS(AVAL).GT.(9.0E+04)) THEN 

                     JERR = -1
                     CALL ERROR (ERRMAT,ERRNUM,6)
                     RETURN

                END IF

                P(I,J) = AVAL

 220       CONTINUE

 250  CONTINUE

      DO 260 I=1,NK

         Q(I) = 0
         R(I) = 0

           DO 260 K=1,NG

                Q(I) = Q(I)+QT(K,I)
                R(I) = R(I)+RT(K,I)

 260  CONTINUE

      END


