!CC******************************************************************
!CC
!CC                                 VBMSCH
!CC                  CALCULATE SCHROEDER'S MOLAR VOLUME
!CC
!CC Output Variablss:
!CC    VBM =      Molar volume at the normal boiling point (cc/gmol)
!CC    MERR =     Error flag
!CC
!CC Input Variables:
!CC    NC =
!CC    IRNG =
!CC
!CC Variables Internal to Subroutine VBMSCH:
!CC    VTM =
!CC    XTS =
!CC
!CC Authors:  M. Miller, T. Rogers, D. Hokanson
!CC
!CC******************************************************************

      SUBROUTINE VBMSCH (NC,VBM,IRNG,MERR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VBMSCH
!MS$ ATTRIBUTES ALIAS:'_VBMSCH@16':: VBMSCH
!MS$ ATTRIBUTES REFERENCE::NC,VBM,IRNG,MERR

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

!CC-----Modified David R. Hokanson 7/9/01 for STEPP2
!CC-----   Increased dimensioning for new binary interaction parameter databases      
!CC      PARAMETER(MA=53,NA=96,ND=10)
      PARAMETER(MA=58,NA=116,ND=10)
!CC-----End Modified David R. Hokanson 7/9/01 for STEPP2

      COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
      COMMON /LIMITS/ TOL,IMAX
      COMMON /GROUP/ MS(10,10,2),NMAX
      COMMON /ERR/ ERRMAT(30),ERRNUM

      NK = NC
      VTM = 0.0D0

      DO 108 KJ=1,NMAX

           IDG = MS(NK,KJ,1)

           IF (IDG.EQ.0) GOTO 108

           XTS = FVB(IDG)

           IF (XTS.LE.0) GOTO 109

           VTM = VTM+XTS*DBLE(MS(NK,KJ,2))

 108  CONTINUE

      VBM = VTM

      IF ((NMAX.EQ.1).AND.(MS(NK,NMAX,2).EQ.1)) GOTO 109

      VBM = (VBM-(DBLE(IRNG)*7.0D0))

 109  IF (VBM.LE.TOL) THEN

           MERR = -1
           CALL ERROR (ERRMAT,ERRNUM,8)
           RETURN

      END IF

      END


