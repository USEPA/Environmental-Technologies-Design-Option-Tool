!C***************************************************************
!CC
!CC                       AIRVISC
!CC
!CC Description:   This subroutine will calculate the air viscosity
!CC                given temperature.
!CC
!CC Output Variable:
!CC    VG =         Viscosity of air (kg/m/sec)
!CC
!CC Input Variable:
!CC    TEMP =       Operating temperature (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE AIRVISC(VG,TEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AIRVISC
!MS$ ATTRIBUTES ALIAS:'_AIRVISC':: AIRVISC
!MS$ ATTRIBUTES REFERENCE::VG,TEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VG,TEMP

         VG = (1.7D-7)*(TEMP**0.818)

      END

!C***************************************************************

