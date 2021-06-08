!C**************************************************************
!CC
!CC                     DBDENS
!CC
!CC Description:  This subroutine will calculate the liquid
!CC               density of the chemical using information
!CC               from the DIPPR 801 database.
!CC
!CC Output Variable:
!CC    DBDEN =    Liquid Density of chemical (kg/m^3)
!CC
!CC Input Variables:
!CC    A,B,C,D =  Coefficients for DIPPR 801 equation
!CC    TT =       Operating temperature (K)
!CC    TEMP =     Operating temperature (C)
!CC    FWT =      Molecular weight of the compound
!CC    TMIN =     Minimum temp. at which DIPPR 801 correlation valid
!CC    TMAX =     Maximum temp. at which DIPPR 801 correlation valid
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!C**************************************************************

      SUBROUTINE DBDENS(DBDEN,A,B,C,D,TT,TMIN,TMAX,FWT,NTRGE)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DBDENS
!MS$ ATTRIBUTES ALIAS:'_DBDENS@40':: DBDENS
!MS$ ATTRIBUTES REFERENCE::DBDEN,A,B,C,D,TT,TMIN,TMAX,FWT,NTRGE

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DBDEN,A,B,C,D,TT,TEMP

         DBDEN = A/(B**(1.0D0+((1.0D0-TT/C)**D)))
         DBDEN = DBDEN*FWT

         TEMP = TT - 273.15D0
         IF ((TMIN.GT.(-273.15D0)).AND.((TEMP.LE.TMIN).OR.(TEMP.GE.TMAX))) THEN
            NTRGE = -1 
         END IF

      END

!C**************************************************************


