!C**************************************************************
!CC
!CC                     ORGDENS
!CC
!CC Description:  This subroutine will estimate the liquid
!CC               density of any organic chemical making use of
!CC               group contribution method with Tony Rogers'
!CC               manipulation of Schroeder's molar volume.
!CC
!CC Output Variable:
!CC    ORGDEN =   Liquid Density (kg/m^3)
!CC
!CC Input Variables:
!CC    FWT =      Molecular weight of the compound
!CC    VBM =      Molar volume at the normal boiling point (m^3/kmol)
!CC    DLH2O =    Density of water at the temperature of interest (kg/m^3)
!CC
!CC Variables Internal to Subroutine ORGDENS:
!CC    PW =
!CC
!CC Author:  D. Hokanson, T. Rogers (4/5/94)
!CC
!C**************************************************************

      SUBROUTINE ORGDENS(ORGDEN,FWT,VBM,DLH2O)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ORGDENS
!MS$ ATTRIBUTES ALIAS:'_ORGDENS@16':: ORGDENS
!MS$ ATTRIBUTES REFERENCE::ORGDEN,FWT,VBM,DLH2O

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION ORGDEN,FWT,VBM,DLH2O

         PW =  0.95D0
         ORGDEN = PW*(DLH2O/1000.0D0)*(FWT/(VBM*1000.0D0)) / (18.015D0/21.D0)
         ORGDEN = ORGDEN * 1000.0D0

      END

!C**************************************************************


