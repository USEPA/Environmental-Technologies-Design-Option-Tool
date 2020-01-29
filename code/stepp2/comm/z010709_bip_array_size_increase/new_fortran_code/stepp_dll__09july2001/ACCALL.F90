!CC************************************************************
!CC
!CC                      ACCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to get activity coefficient values
!CC
!CC Output Variables:
!CC    VAL =      Activity coefficient value (dimensionless)
!CC    SRCSHT =   Source of the value (Short Version)
!CC    SRCLNG =   Source of the value (Long Version)
!CC    ERRORF =   Error Flag
!CC    TEMPUN =   Temperature of the activity coefficient value
!CC    FGRPER =   Error flag from call to subroutine FGRPCALL
!CC
!CC Input Variables:
!CC    TEMPOP =   Operating temperature (C)
!CC    MX =       Maximum number of UNIFAC groups
!CC    INPMS =    Input array of UNIFAC groups in compound
!CC
!CC Variables Internal to Subroutine ACCALL:
!CC    MS =       Array of UNIFAC groups needed to do calculations in
!CC               COMMON.  It holds same values as INPMS, but could not
!CC               be input directly due to use of MS in a COMMON block.
!CC    MDL =      Choice of UNIFAC parameter set:  3=Environmental VLE,
!CC               2=Liquid-Liquid Equilibrium (LLE), or 1=Vapor-Liquid
!CC               Equilibrium (VLE)
!CC    GAMMA =    Activity coefficient value
!CC    NC =
!CC    NG =
!CC    TT =       Operating temperature (K)
!CC    NDIF =
!CC    XX =
!CC    ACT =
!CC    DACT =
!CC    TACT =
!CC
!CC Author:  D. Hokanson (4/3/94)
!CC
!CC************************************************************
       
      SUBROUTINE ACCALL(VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMPOP,FGRPER,MX,INPMS,MDL)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ACCALL
!MS$ ATTRIBUTES ALIAS:'_ACCALL':: ACCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMPOP,FGRPER,MX,INPMS,MDL

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
!CC--------Modified David R. Hokanson 7/9/01 for STEPP2
!CC--------   Increased dimensioning for new binary interaction parameter databases      
!CC         PARAMETER(MA=53,NA=96,NC=2,ND=10)
         PARAMETER(MA=58,NA=116,NC=2,ND=10)
!CC--------End Modified David R. Hokanson 7/9/01 for STEPP2
         COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
         COMMON /GROUP/ MS(10,10,2),NMAX
         COMMON /UNI/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),R(10),P(10,10)
         COMMON /LIMITS/ TOL,IMAX
         COMMON /INIT/ XX(10), NG, NDIF
         DIMENSION ACT(ND),TACT(ND),DACT(10,10)
         
         DOUBLE PRECISION VAL,TEMPUN,TEMPOP
         INTEGER SRCSHT,SRCLNG,ERRORF,FGRPER
         DIMENSION  NGM(10),NY(10,20),JH(NA),IH(20)
         DIMENSION INPMS(10,10,2)

         ERRORF = 0
         SRCSHT = 7

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
!CC      *    Use MDL passed in from Visual Basic Code to        *
!CC      *    choose which database                              *
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

       CALL FGRPCALL(FGRPER)

         IF (FGRPER.EQ.-2) THEN
            ERRORF = -2
         ELSE IF (FGRPER.EQ.-3) THEN
            ERRORF = -3
         ELSE

            TT = TEMPOP + 273.15D0
            CALL GETGAM(GAMMA,NC,NG,TT,NDIF,XX,ACT,DACT,TACT)
            VAL = GAMMA
            TEMPUN = TEMPOP
         
         END IF
      RETURN
      END

!CC************************************************************

