!C****************************************************************
!CC
!CC                    ONDAKLPT
!CC
!CC Description:  This subroutine will calculate liquid phase
!CC               mass transfer coefficient, KL, using the
!CC               Onda correlation.
!CC
!CC Output Variable:
!CC    KL =       Liquid phase mass transfer coefficient (m/sec)
!CC
!CC Input Variables:
!CC    ML =       Liquid mass loading rate (kg/m^2/sec)
!CC    AW =       Wetted surface area of the packing (m^2/m^3)
!CC    VL =       Liquid viscosity (kg/m/sec)
!CC    DL =       Liquid density (kg/m^3)
!CC    DIFL =     Liquid diffusivity (m^2/sec)
!CC    AT =       Specific surface area of the packing (m^2/m^3)
!CC    DP =       Nominal diameter of the packing (m)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C****************************************************************

      SUBROUTINE ONDAKLPT(KL,ML,AW,VL,DL,DIFL,AT,DP)
!C  ATTRIBUTES DLLEXPORT, STDCALL::ONDAKLPT
!C  ATTRIBUTES ALIAS:'_ONDAKLPT@32':: ONDAKLPT
!C  ATTRIBUTES REFERENCE::KL,ML,AW,VL,DL,DIFL,AT,DP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION KL,ML,AW,VL,DL,DIFL,AT,DP

         KL = 0.0051D0 * ((ML/(AW*VL))**(0.6666666)) * ((VL/(DL*DIFL))**(-0.5)) * ((AT*DP)**0.4) / ((DL/(VL*9.81D0))**(0.3333333))
 
      END

!C****************************************************************


