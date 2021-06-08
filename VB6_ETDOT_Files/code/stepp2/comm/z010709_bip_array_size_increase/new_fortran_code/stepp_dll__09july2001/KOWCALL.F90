!CC************************************************************
!CC
!CC                      KOWCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to calculate unifac log octanol water
!CC   partition coefficient (log Kow) at given TEMP
!CC
!CC Output Variables:
!CC    VAL =      log Kow value (-)
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    ERRORF =   Error flag
!CC    TEMPUN =   Temperature of this value (C)
!CC    IFGRPERR = Error flag from call to FGRPCALL
!CC
!CC Input Variables:
!CC    TEMP =     Temperature of calculation (C)
!CC    MX =       Maximum number of UNIFAC groups
!CC    INPMS =    Input array of UNIFAC groups in compound
!CC
!CC Variables Internal to Subroutine KOWCALL:
!CC    MS =       Array of UNIFAC groups needed to do calculations in
!CC               COMMON.  It holds same values as INPMS, but could not
!CC               be input directly due to use of MS in a COMMON block.
!CC    MDL =      Choice of UNIFAC parameter set:  3=Environmental VLE,
!CC               2=Liquid-Liquid Equilibrium (LLE), or 1=Vapor-Liquid
!CC               Equilibrium (VLE)
!CC    TT =       Temperature of calculation (K)
!CC    OCTDEN =
!CC    WATDEN =
!CC    JJERR =    Error Flag from call to PARTC
!CC    XKOW =     Octanol water partition coefficient (-)
!CC    XLGK =     log Kow (-)
!CC    NG =
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!CC************************************************************
       
      SUBROUTINE KOWCALL(VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMP,IFGRPERR,MX,INPMS,MDL)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KOWCALL
!MS$ ATTRIBUTES ALIAS:'_KOWCALL':: KOWCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMP,IFGRPERR,MX,INPMS,MDL

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)

!CC--------Modified David R. Hokanson 7/9/01 for STEPP2
!CC--------   Increased dimensioning for new binary interaction parameter databases      
!CC         PARAMETER(MA=53,NA=96,NC=2,ND=10)
         PARAMETER(MA=58,NA=116,NC=2,ND=10)
!CC--------End Modified David R. Hokanson 7/9/01 for STEPP2
         
         DOUBLE PRECISION VAL,TEMPUN,TEMPOP
         INTEGER SRCSHT,SRCLNG,ERRORF
         COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
         COMMON /GROUP/ MS(10,10,2),NMAX
         COMMON /UNI/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),R(10),P(10,10)
         COMMON /INIT/ XX(10), NG, NDIF   
         COMMON /LIMITS/ TOL,IMAX
         DIMENSION INPMS(10,10,2)


!CC      *********************************************************
!CC      *                                                       *
!CC      *       Set variables needed by COMMONS:  MS            *
!CC      *                                                       *
!CC      *********************************************************

       DO 3, I = 1,10
          DO 3, J=1,10
             DO 3, K=1,2
                MS(I,J,K) = INPMS(I,J,K)
 3     CONTINUE


!CC      *********************************************************
!CC      *                                                       *
!CC      *                Initialize Variables                   *
!CC      *                                                       *
!CC      *********************************************************

       CALL INITVS(MX)  

!CC      *********************************************************
!CC      *                                                       *
!CC      *          Load UNIFAC binary interaction parameters    *
!CC      *    Use MDL supplied from Visual Basic corresponding   *
!CC      *    to user's choice of UNIFAC database                *
!CC      *                                                       *
!CC      *       MDL       Database                              *
!CC      *                                                       *
!CC      *        1 = Original UNIFAC VLE (AVLE.DAT)             *
!CC      *        2 = UNIFAC LLE (ALLE.DAT)                      *
!CC      *        3 = Environmental VLE (AENV.DAT)               *
!CC      *                                                       *
!CC      *********************************************************

       CALL BINPAR(MDL,MGSG,AI,RI,QI,FMW,FVB)


!CC      ******************************************************
!CC      *                                                    *
!CC      *           Reorder the functional groups            *
!CC      *                                                    *
!CC      ******************************************************

       CALL FGRPCALL(IFGRPERR)

         IF (IFGRPERR.EQ.-2) THEN
            ERRORF = -2
         ELSE IF (IFGRPERR.EQ.-3) THEN
            ERRORF = -3
         ELSE
            ERRORF = 0
            SRCSHT = 7
            TT = TEMP + 273.15D0
            OCTDEN = 6.36D0
            WATDEN = 55.5D0
            CALL PARTC(TT,OCTDEN,WATDEN,XKOW,XLGK,NG,JJERR)

!CC           ***** Set functional groups back to there original ordering

            CALL FGRPCALL(IFGRPERR)
  
            IF (JJERR.EQ.-1) THEN
               ERRORF = -23
            ELSE IF (JJERR.EQ.-2) THEN
               ERRORF = -24
            ELSE IF (JJERR.EQ.-3) THEN
               ERRORF = -25
            ELSE IF (JJERR.EQ.-4) THEN
               ERRORF = -26
            ELSE IF (JJERR.EQ.-5) THEN
               ERRORF = -27
            ELSE IF (JJERR.EQ.-6) THEN
               ERRORF = -28
            ELSE IF (JJERR.EQ.-7) THEN
               ERRORF = -29
            ELSE IF (JJERR.EQ.-8) THEN
               ERRORF = -30
            ELSE
               VAL = XLGK
               TEMPUN = TEMP
            END IF
         END IF
      END

!CC************************************************************

