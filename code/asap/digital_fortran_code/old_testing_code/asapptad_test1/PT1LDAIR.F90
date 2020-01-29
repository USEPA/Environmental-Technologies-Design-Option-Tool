!C***************************************************************
!CC
!CC                       PT1LDAIR
!CC
!CC Description:  This subroutine will calculate air mass loading
!CC               rate from Figure 6.34 of Treybol, 1980.
!CC
!CC Output Variables:
!CC    GM =       Air mass loading rate (kg/m^2/sec)
!CC
!CC Input Variables:
!CC    PRESD =    Desired air pressure drop across tower (N/m2/m)
!CC    VQ =       Air to water ratio (dimensionless)
!CC    DG =       Air density (kg/m^3)
!CC    DL =       Water density (kg/m^3)
!CC    CF =       Packing factor (dimensionless)
!CC    VL =       Liquid viscosity (kg/m/sec)
!CC
!CC Variables Internal to subroutine LDAIRPT1
!CC    FF,A0,A1 = Temporary values used to calculate air
!CC    A2,EE,MM   mass loading rate, GM (kg/m2-sec)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PT1LDAIR(GM,PRESD,VQ,DG,DL,CF,VL)     
!C  ATTRIBUTES DLLEXPORT, STDCALL::PT1LDAIR
!C  ATTRIBUTES ALIAS:'_PT1LDAIR@28':: PT1LDAIR
!C  ATTRIBUTES REFERENCE::GM,PRESD,VQ,DG,DL,CF,VL

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION GM,PRESD,FF,A0,A1,A2,EE,MM,VQ,DG,DL,CF,VL

      FF = LOG10(PRESD)                                  
      A0 = -6.6599D0+4.3077D0*FF-1.3503D0*(FF**2)+0.15931D0*(FF**3)         
      A1 = 3.0945D0-4.3512D0*FF+1.6240D0*(FF**2)-0.20855D0*(FF**3)        
      A2 = 1.7611D0-2.3394D0*FF+0.89914D0*(FF**2)-0.11597D0*(FF**3)      
      EE = -LOG10(VQ*(((DG/DL)-((DG/DL)**2))**0.5))
      MM =  10**(A0 + A1*EE + A2*(EE**2))
      GM = ((MM*DG*(DL-DG))/(CF*(VL**0.1)))**0.5               

      END

!C**************************************************************

