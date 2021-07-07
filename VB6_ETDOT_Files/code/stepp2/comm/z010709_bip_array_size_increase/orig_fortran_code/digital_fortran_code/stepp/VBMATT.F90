!CC*********************************************************************
!CC
!CC                           VBMATT
!CC         CALCULATE MOLAR VOLUME AS INVERSE OF LIQUID DENSITY
!CC
!CC Description:  This subroutine will calculate molar volume at the
!CC               temperature of interest by taking the inverse of
!CC               the liquid density at the temp. of interest.
!CC
!CC Output Variable:
!CC    VBMTMP =   Molar volume at the temperature of interest (m^3/kmol)
!CC
!CC Input Variables:
!CC    LIQDEN =   Liquid density of chemical at temp. of interest (kg/m^3)
!CC    FWT =      Molecular weight of the compound
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!CC*********************************************************************

      SUBROUTINE VBMATT(VBMTMP,LIQDEN,FWT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VBMATT
!MS$ ATTRIBUTES ALIAS:'_VBMATT':: VBMATT
!MS$ ATTRIBUTES REFERENCE::VBMTMP,LIQDEN,FWT

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION VBMTMP,LIQDEN,FWT

         VBMTMP = (1.0D0/LIQDEN)*FWT

      RETURN
      END

!CC********************************************************************

