!C***************************************************************
!CC
!CC                       PDROP
!CC
!CC Description:  This subroutine will calculate the pressure
!CC               drop from a routine developed from the Eckhert
!CC               curve (from Cummings?).  Note:  the Eckhert curve
!CC               is only valid in approximate pressure drop range
!CC               of 50 - 1200 N/m^2/m and operating above a
!CC               pressure drop of 300 N/m^2/m is generally not
!CC               desirable in practice.
!CC
!CC Output Variables:
!CC    PRESD =    Gas pressure drop (N/m^2/m)
!CC
!CC Input Variables:
!CC    VQ =       Air to water ratio (dimensionless)
!CC    GM =       Air mass loading rate (kg/m^2/sec)`
!CC    CF =       Packing factor (dimensionless)
!CC    VL =       Water viscosity (kg/m-sec)
!CC    DG =       Air density (kg/m^3)
!CC    DL =       Water density (kg/m^3)
!CC    DEL =      Iterative step used to find Pressure drop
!CC
!CC Variables internal to subroutine PDROP:
!CC    YYA =      Value of the y-axis on the Eckert Curve
!CC    PP =       Iterative pressure drop
!CC    DRPINI =   Initial pressure drop used in iteration
!CC               (set equal to DEL right now)
!CC    DELTA =    Step to use in iteration
!CC               (set equal to 1 right now)
!CC    DRPMAX =   Maximum pressure drop used in iteration
!CC               (set equal to 1200 right now)
!CC  EE,PD,A0,A1,A2 = Values used in subroutine to make calculation
!CC    YYB =      Calculated value on y-axis of Eckert curve
!CC               based on the current pressure drop.
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PDROP(PRESD,VQ,GM,CF,VL,DG,DL,DRPINI,DRPMAX,DEL)
!C  ATTRIBUTES DLLEXPORT, STDCALL::PDROP
!C  ATTRIBUTES ALIAS:'_PDROP':: PDROP
!C  ATTRIBUTES REFERENCE::PRESD,VQ,GM,CF,VL,DG,DL,DRPINI,DRPMAX,DEL

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VQ,GM,CF,VL,DG,DL,YYA,PP,DRPINI,DELTA,DRPMAX,EE,PD,A0,A1,A2,YYB,DEL

         YYA = ((GM**2)*CF*(VL**0.1D0))/(DG*(DL-DG))
         DRPINI = 1.0D0
         PP = DRPINI
         DELTA = DEL
         DRPMAX = 1200.0D0
         EE = -1.0D0*LOG10(VQ*(((DG/DL)-(DG/DL)**2)**0.5))
1300     PD = LOG10(PP)
         A0 = -6.6599D0 + 4.3077D0*PD - 1.3503D0*PD**2 + 0.15931D0*PD**3
         A1 = 3.0945D0 - 4.3512D0*PD + 1.6240D0*PD**2 - 0.20855D0*PD**3
         A2 =  1.7611D0 - 2.3394D0*PD + 0.89914D0*PD**2 - 0.11597D0*PD**3
         YYB = 10.0D0**(A0+A1*EE+A2*EE**2)

         IF (PP.GT.DRPMAX) THEN
            PRESD = -1
         ELSE IF ((YYB.LT.(0.99D0*YYA)).OR.(YYB.GT.(1.010D0*YYA))) THEN
            PP = PP + DELTA
            GOTO 1300
         ELSE
            PRESD = PP
         END IF

      END

!C***************************************************************
  
