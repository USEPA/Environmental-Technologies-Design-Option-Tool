!C*****************************************************************
!CC
!CC                          H2OST
!CC
!CC Description:  This subroutine will calculate the surface
!CC               tension of water, given temperature.  The
!CC               correlation comes from a routine presented by
!CC               Cummins and Westrick (1983)
!CC
!CC Output Variables:
!CC    ST =       Value of surface tension of water (kg/sec^2)
!CC    ERRORF =   Error flag
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    STTEMP =   Temperature of this value (C)
!CC
!CC Input Variable:
!CC    TEMPOP =   Temperature of this calculation (C)
!CC
!CC Variables Internal to Subroutine H2OST:
!CC    TEMP =     Temperature of the calculation (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C*****************************************************************

      SUBROUTINE H2OST(ST,TEMPOP,ERRORF,SRCSHT,SRCLNG,STTEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::H2OST
!MS$ ATTRIBUTES ALIAS:'_H2OST':: H2OST
!MS$ ATTRIBUTES REFERENCE::ST,TEMPOP,ERRORF,SRCSHT,SRCLNG,STTEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION ST,TEMP
      INTEGER ERRORF,SRCSHT,SRCLNG

         ERRORF = 0
         SRCSHT = 15
         STTEMP = TEMPOP
         TEMP = TEMPOP + 273.15D0
         ST = (7.558301D-2) - (1.3143D-4)*(TEMP-273.15D0) - (4.7616D-7)*((TEMP-273.15D0)**2)

      END

!C*****************************************************************

                      
