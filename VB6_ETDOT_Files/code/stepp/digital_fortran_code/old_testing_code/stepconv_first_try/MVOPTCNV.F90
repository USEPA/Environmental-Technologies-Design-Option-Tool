!CC********************************************************************
!CC
!CC                             MVOPTCNV
!CC          CONVERT MOLAR VOLUME AT OPERATING TEMP FROM M3/KMOL TO
!CC          FT3/LBm-MOL
!CC
!CC Description:  This SUBROUTINE will convert molar volume at
!CC               operating temp units from m3/Kmol to Ft3/LBm-mol.
!CC
!CC Output Variables:
!CC    MVOENG =    Molar Volume at Operating Temp (Ft3/LBm-mol)
!CC
!CC Input Variables:
!CC    MVOSI =     Molar Volume at Operating Temp (m3/Kmol)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE MVOPTCNV(MVOENG,MVOSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVOPTCNV
!MS$ ATTRIBUTES ALIAS:'_MVOPTCNV'::MVOPTCNV
!MS$ ATTRIBUTES REFERENCE::MVOENG,MVOSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION MVOENG, MVOSI

         MVOENG = MVOSI * 35.3145D0/2.20462D0           

      END

!CC********************************************************************


