!C**************************************************************
!CC
!CC                     H20DENS
!CC
!CC Description:  This subroutine will estimate the liquid
!CC               density using a routine from Reid, Prausnitz,
!CC               and Poling (1987).
!CC
!CC Output Variable:
!CC    DL =       Liquid Density (kg/m^3)
!CC
!CC Input Variables:
!CC    TEMP =     Temperature in Deg K
!CC
!CC Variables Internal to Subroutine H2ODENS:
!CC     A1,A2,A3,A4,A5 = Constants used in polynomial fit
!CC     XAVG,FAVG,XN,FN = Variables used to calculate DL
!CC     FX = Liquid Density of Water (g/cm^3)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C**************************************************************

      SUBROUTINE H2ODENS(DL,TEMP)
!C  ATTRIBUTES DLLEXPORT, STDCALL::H2ODENS
!C  ATTRIBUTES ALIAS:'_H2ODENS':: H2ODENS
!C  ATTRIBUTES REFERENCE::DL,TEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION A1,A2,A3,A4,A5,XAVG,FAVG,XN,TEMP,FN,FX,DL
      
         A1 = -1.4176800403D+00
         A2 =  8.9766515240D+00
         A3 = -1.2275501969D+01
         A4 =  7.4584410413D+00
         A5 = -1.7384916050D+00
         XAVG = 324.65D+00    
         FAVG = 0.98396D+00  
         XN = TEMP/XAVG       
         FN = A1 + A2*(XN) + A3*(XN)**2 + A4*(XN)**3 + A5*(XN)**4 

         FX = FN*FAVG                                            

         DL = FX * 1000.0D0

      END

!C**************************************************************

