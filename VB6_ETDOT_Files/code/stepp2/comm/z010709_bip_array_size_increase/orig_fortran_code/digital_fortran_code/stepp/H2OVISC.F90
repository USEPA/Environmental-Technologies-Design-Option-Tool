!C*****************************************************************
!CC
!CC                      H2OVISC
!CC
!CC Description:  This subroutine will calculate the viscosity of
!CC               liquid water using a routine from Yaws, et. al. (1976)
!CC
!CC Output Variable:
!CC    VL =       Water viscosity value (kg/m/sec)
!CC    ERRORF =   Error Flag
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    VLTEMP =   Temperature of this value (C)
!CC
!CC Input Variables:
!CC    TEMPOP =   Temperature of the calculation (C)
!CC
!CC Variables Internal to Subroutine H2OVISC:
!CC    TEMP =     Temperature of the calculation (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C*****************************************************************

      SUBROUTINE H2OVISC(VL,TEMPOP,ERRORF,SRCSHT,SRCLNG,VLTEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::H2OVISC
!MS$ ATTRIBUTES ALIAS:'_H2OVISC@24':: H2OVISC
!MS$ ATTRIBUTES REFERENCE::VL,TEMPOP,ERRORF,SRCSHT,SRCLNG,VLTEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VL,TEMP,TEMPOP
      INTEGER ERRORF,SRCSHT,SRCLNG

         ERRORF = 0
         SRCSHT = 1
         VLTEMP = TEMPOP
         TEMP = TEMPOP + 273.15D0
         VL = EXP(-24.71D0+(4209.0D0/TEMP)+(.04527D0*TEMP) - (3.376D-5) * (TEMP**2))
         VL = VL/1000.0D0

      END

!C*****************************************************************


