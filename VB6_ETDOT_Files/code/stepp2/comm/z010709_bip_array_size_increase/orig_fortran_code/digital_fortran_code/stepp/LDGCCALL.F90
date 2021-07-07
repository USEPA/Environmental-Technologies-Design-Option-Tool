!CC************************************************************
!CC
!CC                      LDGCCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to get liquid density values
!CC   from the Group Contribution Method related to Schroeder's
!CC   method.
!CC
!CC Output Variables:
!CC    VAL =       Liquid density value (kg/m3)
!CC    SRCSHT =    Source of this value (Short Version)
!CC    SRCLNG =    Source of this value (Long Version)
!CC    ERRORF =    Error flag
!CC    TEMPUN =    Temperature of this value (C)
!CC
!CC Input Variables:
!CC    FWT =       Molecular weight (kg/kmol)
!CC    VBMNBP =    Molar volume at the normal boiling point (m3/kmol)
!CC    TEMPOP =    Operating temperature (C)
!CC
!CC Variables Internal to Subroutine LDGCALL:
!CC    DLH2O =     Density of water (kg/m3)
!CC    MERR =      Error flag from water density calculation
!CC    IDUMSHT =   Dummy source variable for water density calculation
!CC    IDUMLNG =   Dummy source variable for water density calculation
!CC    DUMTMP =    Dummy temp. variable for water density calculation
!CC    ORGDEN =    Liquid density of the chemical (kg/m3)
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!CC************************************************************
       
      SUBROUTINE LDGCCALL(VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,FWT,VBMNBP,TEMPOP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDGCCALL
!MS$ ATTRIBUTES ALIAS:'_LDGCCALL':: LDGCCALL
!MS$ ATTRIBUTES REFERENCE::VAL,SRCSHT,SRCLNG,ERRORF,TEMPUN,FWT,VBMNBP,TEMPOP

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)
         DOUBLE PRECISION VAL,TEMPUN,TEMPOP,FWT,VBMNBP
         INTEGER SRCSHT,SRCLNG,ERRORF

         ERRORF = 0
         SRCSHT = 9

         TEMPUN = TEMPOP

         CALL H2ODENS(DLH2O,TEMPUN,MERR,IDUMSHT,IDUMLNG,DUMTMP)
         CALL ORGDENS(ORGDEN,FWT,VBMNBP,DLH2O)
         VAL = ORGDEN
         
      END

!CC************************************************************


