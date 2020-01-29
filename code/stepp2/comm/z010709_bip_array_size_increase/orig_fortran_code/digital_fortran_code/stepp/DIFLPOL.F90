!C*************************************************************
!CC
!CC                    DIFLPOL
!CC
!CC Description:  This subroutine will calculate liquid diffusivity
!CC               for compounds.  It is generally valid for
!CC               molecular weight > 1000.  It uses the method of
!CC               Polson, 1950.
!CC
!CC Output Variables:
!CC    DIFL =     Liquid diffusivity value (m^2/sec)
!CC    ERRORF =   Error flag
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    DIFLT =    Temperature of this calculation (C)
!CC
!CC Input Variables:
!CC    MW =       Molecular weight of compound
!CC    TEMPOP     Temperature of this calculation (C)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C*************************************************************

      SUBROUTINE DIFLPOL(DIFL,MW,ERRORF,SRCSHT,SRCLNG,DIFLT,TEMPOP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DIFLPOL
!MS$ ATTRIBUTES ALIAS:'_DIFLPOL':: DIFLPOL
!MS$ ATTRIBUTES REFERENCE::DIFL,MW,ERRORF,SRCSHT,SRCLNG,DIFLT,TEMPOP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFL,MW
      INTEGER ERRORF,SRCSHT,SRCLNG

         ERRORF = 0
         SRCSHT = 11
         DIFLT = TEMPOP
         DIFL = (2.74D-5)*(MW**(-1.0D0/3.0D0))            
!C******** CHANGE MADE ON 08-FEB-1999 BEGINS:
         DIFL = DIFL / 10000.0D0
!C******** CHANGE MADE ON 08-FEB-1999 ENDS.
         IF (MW.LT.1000) THEN
            ERRORF = 4
         END IF
      END

!C*************************************************************


