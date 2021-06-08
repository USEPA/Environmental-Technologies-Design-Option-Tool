!CC***************************************************************************
!CC
!CC                                  MOLWT
!CC                     CALCULATE UNIFAC MOLECULAR WEIGHT
!CC
!CC Output Variables:
!CC    FWT =      Molecular weight value (kg/kmol)
!CC    XMW =
!CC    NERR =     Error flag
!CC
!CC Input Variables:
!CC    NC =
!CC
!CC Authors:      M. Miller, T. Rogers, D. Hokanson
!CC
!CC***************************************************************************

      SUBROUTINE MOLWT (FWT,NC,XMW,NERR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MOLWT
!MS$ ATTRIBUTES ALIAS:'_MOLWT@16':: MOLWT
!MS$ ATTRIBUTES REFERENCE::FWT,NC,XMW,NERR

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

!CC-----Modified David R. Hokanson 7/9/01 for STEPP2
!CC-----   Increased dimensioning for new binary interaction parameter databases      
!CC      PARAMETER (MA=53,NA=96,ND=10)
      PARAMETER(MA=58,NA=116,ND=10)
!CC-----End Modified David R. Hokanson 7/9/01 for STEPP2

      DIMENSION XMW(ND),XTW(ND)

      COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
      COMMON /LIMITS/ TOL,IMAX
      COMMON /GROUP/ MS(10,10,2),NMAX
      COMMON /ERR/ ERRMAT(30),ERRNUM

      NK = NC

      DO 103 KO=1,3

           XMW(KO) = 0.0D0

 103  CONTINUE

      XMW(1) = 18.015D0
      XMW(NK) = FWT

      DO 107 LP=1,NK

           XTW(LP) = 0.0D0

           DO 105 JP=1,NMAX

                IDG = MS(LP,JP,1)

                IF (IDG.EQ.0) GOTO 105

                XTS = FMW(IDG)

                IF (XTS.LE.0.D0) GOTO 107

                XTW(LP) = XTW(LP)+XTS*DBLE(MS(LP,JP,2))

 105       CONTINUE

           IF(XTW(LP).GT.TOL) XMW(LP)=XTW(LP)

 107  CONTINUE

      FWT = XMW(NK)

 109  IF (FWT.LE.TOL) THEN

           NERR = -1
           CALL ERROR (ERRMAT,ERRNUM,7)
           RETURN

      END IF
 
      END


