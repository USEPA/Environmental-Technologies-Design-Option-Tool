!CC********************************************************************
!CC
!CC                             MVNBPCNV
!CC          CONVERT MOLAR VOLUME AT NORMAL BOILING TEMP FROM M3/KMOL TO
!CC          FT3/LBm-MOL
!CC
!CC Description:  This SUBROUTINE will convert molar volume at
!CC               normal boiling point units from m3/Kmol to Ft3/LBm-mol.
!CC
!CC Output Variables:
!CC    MVNENG =    Molar Volume at Normal Boiling Pt (Ft3/LBm-mol)
!CC
!CC Input Variables:
!CC    MVNSI =     Molar Volume at Normal Boiling Pt (m3/Kmol)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE MVNBPCNV(MVNENG,MVNSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVNBPCNV
!MS$ ATTRIBUTES ALIAS:'_MVNBPCNV'::MVNBPCNV
!MS$ ATTRIBUTES REFERENCE::MVNENG,MVNSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION MVNENG, MVNSI

         MVNENG = MVNSI * 35.3145D0/2.20462D0           

      END

!CC********************************************************************


