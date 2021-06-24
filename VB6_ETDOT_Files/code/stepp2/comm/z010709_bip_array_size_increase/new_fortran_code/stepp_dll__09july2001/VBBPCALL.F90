!CC************************************************************
!CC
!CC                      VBBPCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutine to get molar volume at normal boiling point
!CC
!CC Output Variables:
!CC    VAL =      Molar volume at the normal boiling point (m3/kmol)
!CC    SRCSHT =   Source of the value (Short Version)
!CC    SRCLNG =   Source of the value (Long Version)
!CC    ERRORF =   Error flag
!CC    TEMPSC =   Temp. of molar volume at NBP = NBP (C)
!CC
!CC Input Variables:
!CC    TEMPBP =   Boiling point temperature (C)
!CC    MX =       Maximum number of UNIFAC groups
!CC    INPMS =    Input array of UNIFAC groups in compound
!CC    IRNG =     Number of rings in compound
!CC
!CC Variables Internal to Subroutine VBBPCALL:
!CC
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!CC************************************************************
       
      SUBROUTINE VBBPCALL(VAL,SRCSHT,SRCLNG,ERRORF,TEMPSC,TEMPBP,MX,INPMS,IRNG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VBBPCALL
!MS$ ATTRIBUTES ALIAS:'_VBBPCALL':: VBBPCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,TEMPSC,TEMPBP,MX,INPMS,IRNG
      
         PARAMETER  (NC = 2)

!CC--------Modified David R. Hokanson 7/9/01 for STEPP2
!CC--------   Increased dimensioning for new binary interaction parameter databases      
!CC         PARAMETER(MA=53,NA=96,ND=10)
         PARAMETER(MA=58,NA=116,ND=10)
!CC--------End Modified David R. Hokanson 7/9/01 for STEPP2
         
         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
         DIMENSION INPMS(10,10,2)
         DOUBLE PRECISION VAL,TEMPSC,TEMPBP
         INTEGER SRCSHT,ERRORF,SRCLNG
      COMMON /ACTCO/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
      COMMON /LIMITS/ TOL,IMAX
      COMMON /GROUP/ MS(10,10,2),NMAX

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
         SRCSHT = 8
         TEMPSC = TEMPBP
         CALL VBMSCH(NC,VBM,IRNG,MERR)
         VBM = VBM / 1000.0D0           
         IF (NERR.EQ.-1) THEN
            ERRORF = -7
         ELSE
            VAL = VBM
         END IF
           
      END

!CC************************************************************

