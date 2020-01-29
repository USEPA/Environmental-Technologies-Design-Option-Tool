!C**************************************************************
!CC
!CC                     H2ODENS
!CC
!CC Description:  This subroutine will estimate the liquid
!CC               density using a routine created by Tony Rogers.
!CC
!CC Output Variable:
!CC    DL =       Liquid Density value (kg/m^3)
!CC    ERRORF =   Error flag
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    DLTEMP =   Temperature of the calculation (C)
!CC
!CC Input Variable:
!CC    TEMPOP =   Temperature of the Calculation (C)
!CC
!CC Variables Internal to Subroutine H2ODENS:
!CC     A1,A2,A3,A4,A5 = Constants used in polynomial fit
!CC     XAVG,FAVG,XN,FN = Variables used to calculate DL
!CC     FX = Liquid Density of Water (g/cm^3)
!CC     TEMP =     Temperature of the calculation (K)
!CC
!CC Author:  T. Rogers and D. Hokanson (4/5/94)
!CC
!C**************************************************************

	  SUBROUTINE H2ODENS(DL,TEMPOP,ERRORF,SRCSHT,SRCLNG,DLTEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::H2ODENS
!MS$ ATTRIBUTES ALIAS:'_H2ODENS@24':: H2ODENS
!MS$ ATTRIBUTES REFERENCE::DL,TEMPOP,ERRORF,SRCSHT,SRCLNG,DLTEMP

	  IMPLICIT DOUBLE PRECISION(A-H,O-Z)
	  DOUBLE PRECISION A1,A2,A3,A4,A5,XAVG,FAVG,XN,TEMP,FN,FX,DL
	  INTEGER ERRORF,SRCSHT,SRCLNG
	  
		 ERRORF = 0
                 SRCSHT = 14
                 DLTEMP = TEMPOP
		 TEMP = TEMPOP + 273.15D0
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


