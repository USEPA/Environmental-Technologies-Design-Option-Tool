!C*****************************************************************
!CC
!CC                      H2OVISC
!CC
!CC Description:  This subroutine will calculate the viscosity of
!CC               liquid water using a routine from Reid, Prausnitz,
!CC               and Poling (1987).
!CC
!CC Output Variable:
!CC    VL =       Viscosity of liquid water (kg/m/sec)
!CC
!CC Input Variables:
!CC    TEMP =     Operating Temperature (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C*****************************************************************

      SUBROUTINE H2OVISC(VL,TEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::H2OVISC
!MS$ ATTRIBUTES ALIAS:'_H2OVISC':: H2OVISC
!MS$ ATTRIBUTES REFERENCE::VL,TEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VL,TEMP

         VL = EXP(-24.71D0 + (4209.0D0 / TEMP) + (.04527D0 * TEMP) - (3.376D-5) * (TEMP**2))
         VL = VL/1000.0D0

      END

!C*****************************************************************

