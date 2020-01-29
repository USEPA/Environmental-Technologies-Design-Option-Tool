!CC************************************************************
!CC
!CC                      MWTCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutine to get molecular weight value
!CC
!CC Output Variables:
!CC    VAL =      Molecular weight value (kg/kmol)
!CC    SRCSHT =   Source of the value (Short Version)
!CC    SRCLNG =   Source of the value (Long Version)
!CC    ERRORF =   Error flag
!CC    XMW =
!CC
!CC Input Variables:
!CC    MX =       Maximum number of UNIFAC groups
!CC    INPMS =    Input array of UNIFAC groups in compound
!CC
!CC Variables Internal to Subroutine MWTCALL:
!CC    MS =       Array of UNIFAC groups needed to do calculations in
!CC               COMMON.  It holds same values as INPMS, but could not
!CC               be input directly due to use of MS in a COMMON block.
!CC    MDL =      Choice of UNIFAC parameter set:  3=Environmental VLE,
!CC               2=Liquid-Liquid Equilibrium (LLE), or 1=Vapor-Liquid
!CC               Equilibrium (VLE)
!CC
!CC Author:  D. Hokanson (4/4/94)
!CC
!CC************************************************************
       
      SUBROUTINE MWTCALL(VAL,SRCSHT,SRCLNG,ERRORF,MX,INPMS,XMW)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MWTCALL
!MS$ ATTRIBUTES ALIAS:'_MWTCALL':: MWTCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,MX,INPMS,XMW
      
         PARAMETER  (MA=53, NA=96, ND=10)
         PARAMETER  (NC = 2)
         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
         DOUBLE PRECISION VAL
         INTEGER SRCSHT,ERRORF,SRCLNG,FGRPER
         COMMON /LIMITS/ TOL,IMAX
         COMMON /GROUP/ MS(10,10,2), NMAX
         COMMON /INIT/ XX(10), NG, NDIF
         COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
      COMMON /UNI/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),R(10),P(10,10)
         DIMENSION INPMS(10,10,2), XMW(ND)



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
!CC      *    Use MDL = 3 --> Environmental VLE for Henry's      *
!CC      *                    Constant calculation               *
!CC      *                                                       *
!CC      *********************************************************

       MDL = 3
       CALL BINPAR(MDL,MGSG,AI,RI,QI,FMW,FVB)


            ERRORF = 0
            SRCSHT = 7
            FWT = 0.0D0
            CALL MOLWT(FWT,NC,XMW,NERR)     
            IF (NERR.EQ.-1) THEN
               ERRORF = -6
            ELSE
               VAL = FWT
            END IF
           
      END

!CC************************************************************


