!C****************************************************************
!CC
!CC                    ONDAKGPT
!CC
!CC Description:  This subroutine will calculate gas phase
!CC               mass transfer coefficient, KG, using the
!CC               Onda correlation.
!CC
!CC Output Variable:
!CC    KG =       Gas phase mass transfer coefficient (m/sec)
!CC
!CC Input Variables:
!CC    GM =       Gas mass loading rate (kg/m^2/sec)
!CC    AT =       Specific surface area of the packing (m^2/m^3)
!CC    VG =       Air viscosity (kg/m/sec)
!CC    DG =       Gas density (kg/m^3)
!CC    DIFG =     Gas diffusivity (m^2/sec)
!CC    DP =       Nominal diameter of the packing (m)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C****************************************************************

      SUBROUTINE ONDAKGPT(KG,GM,AT,VG,DG,DIFG,DP)
!C  ATTRIBUTES DLLEXPORT, STDCALL::ONDAKGPT
!C  ATTRIBUTES ALIAS:'_ONDAKGPT@28':: ONDAKGPT
!C  ATTRIBUTES REFERENCE::KG,GM,AT,VG,DG,DIFG,DP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION KG,GM,AT,VG,DG,DIFG,DP

         KG = 5.23D0 * ((GM/(AT*VG))**0.7) * ((VG/(DG*DIFG))**(0.3333333D0)) * ((AT*DP)**(-2)) * AT * DIFG        
 
      END

!C****************************************************************

