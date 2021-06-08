!CC************************************************************
!CC
!CC                      FGRPCALL
!CC
!CC   Fortran subroutine to handle calling the appropriate
!CC   Fortran subroutines to order the UNIFAC functional
!CC   groups
!CC
!CC Output variables:
!CC    ERRORF =    Error flag
!CC
!CC Variables Internal to Subroutine FGRPCALL:
!CC    JERR =      Error flag returned from subroutine FGRP
!CC    NC =
!CC    NG =
!CC
!CC************************************************************
       
      SUBROUTINE FGRPCALL(ERRORF)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::FGRPCALL
!MS$ ATTRIBUTES ALIAS:'_FGRPCALL@4':: FGRPCALL
!MS$ ATTRIBUTES REFERENCE::ERRORF

         IMPLICIT DOUBLE PRECISION (A-H,O-Z)      
         
         PARAMETER (NC=2,ND=10)
         COMMON /INIT/ XX(10), NG, NDIF     
         COMMON /ERR/ ERRMAT(30),ERRNUM
         INTEGER ERRORF

         ERRORF = 0
         JERR = 0
         CALL FGRP(NC,NG,JERR)

         IF (JERR.EQ.-1) THEN
            ERRORF = -2
         ELSE IF (JERR.EQ.-2) THEN
            ERRORF = -3
         END IF  
      END

!CC************************************************************

