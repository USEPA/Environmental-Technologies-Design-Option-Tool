!C***************************************************************
!CC
!CC                       AIRVISC
!CC
!CC Description:   This subroutine will calculate the air viscosity
!CC                given temperature.  The correlation comes from
!CC                a paper presented by Cummins and Westrick (1983)
!CC
!CC Output Variables:
!CC    VG =         Air viscosity value (kg/m/sec)
!CC    ERRORF =     Error flag
!CC    SRCSHT =     Source of this value (Short Version)
!CC    SRCLNG =     Source of this value (Long Version)
!CC    VGTEMP =     Temperature of this value
!CC
!CC Input Variables:
!CC    TEMPOP =     Temperature of calculation (C)
!CC
!CC Variables Internal to Subroutine AIRVISC:
!CC    TEMP =       Temperature of the calculation (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C***************************************************************

      SUBROUTINE AIRVISC(VG,TEMPOP,ERRORF,SRCSHT,SRCLNG,VGTEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AIRVISC
!MS$ ATTRIBUTES ALIAS:'_AIRVISC':: AIRVISC
!MS$ ATTRIBUTES REFERENCE::VG,TEMPOP,ERRORF,SRCSHT,SRCLNG,VGTEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VG,TEMP
      INTEGER ERRORF,SRCSHT,SRCLNG

         ERRORF = 0
         SRCSHT = 15
         VGTEMP = TEMPOP
         TEMP = TEMPOP + 273.15D0
         VG = (1.7D-7)*(TEMP**0.818)

      END

!C***************************************************************


