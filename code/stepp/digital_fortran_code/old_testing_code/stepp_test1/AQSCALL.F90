!CC************************************************************
!CC
!CC                      AQSCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to unifac solubility at given TEMP
!CC
!CC Output Variables:
!CC    VAL =      Aqueous solubility value (PPMw)
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    ERRORF =   Error flag
!CC    TEMPUN =   Temeprature of this value (C)
!CC
!CC Input Variables:
!CC    TEMP =     Temperature of calculation (C)
!CC    MX =       Maximum number of UNIFAC groups
!CC    INPMS =    Input array of UNIFAC groups in compound
!CC    XMW =
!CC
!CC Variables Internal to Subroutine AQSCALL:
!CC    MS =       Array of UNIFAC groups needed to do calculations in
!CC               COMMON.  It holds same values as INPMS, but could not
!CC               be input directly due to use of MS in a COMMON block.
!CC    MDL =      Choice of UNIFAC parameter set:  3=Environmental VLE,
!CC               2=Liquid-Liquid Equilibrium (LLE), or 1=Vapor-Liquid
!CC               Equilibrium (VLE)
!CC    FGRPER =   Error flag for call to FGRPCALL
!CC    TT =       Temperature of this calculation (K)
!CC    IIERR =    Error flag from call to AQSOL
!CC    SOLUB =    Solubility of organic in the water phase (PPMw)
!CC    TIE =      Solubility of water in the organic phase (PPMw)
!CC    NC =
!CC    NG =
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!CC************************************************************
       
      SUBROUTINE AQSCALL(VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMP,MX,INPMS,XMW,MDL)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AQSCALL
!MS$ ATTRIBUTES ALIAS:'_AQSCALL':: AQSCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,TEMP,MX,INPMS,XMW,MDL

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
         PARAMETER(MA=53,NA=96,NC=2,ND=10)
         DOUBLE PRECISION VAL,TEMPUN,TEMP
         INTEGER SRCSHT,SRCLNG,ERRORF,FGRPER
         COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
         COMMON /GROUP/ MS(10,10,2),NMAX
         COMMON /UNI/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),R(10),P(10,10)
         COMMON /INIT/ XX(10), NG, NDIF   
         COMMON /LIMITS/ TOL,IMAX
         DIMENSION INPMS(10,10,2),XMW(ND)            


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
!CC      *    Use MDL supplied from Visual Basic to correspond   *
!CC      *    to currently selected UNIFAC parameter set         *
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

            ERRORF = 0
            TT = TEMP + 273.15D0
            CALL AQSOL(NC,NG,TT,SOLUB,TIE,IIERR,XMW)
            IF (IIERR.EQ.-1) THEN
               ERRORF = -18
            ELSE IF (IIERR.EQ.-2) THEN
               ERRORF = -19
            ELSE IF (IIERR.EQ.-3) THEN
               ERRORF = -20
            ELSE IF (IIERR.EQ.-4) THEN
               ERRORF = -21
            ELSE  
               TEMPUN = TEMP
               VAL = SOLUB
               SRCSHT = 7
            END IF


         END IF

      END

!CC************************************************************



