!CC********************************************************************
!CC
!CC                             LDIFFCNV
!CC          CONVERT LIQUID DIFFUSIVITY UNITS FROM m2/sec TO ft2/sec
!CC
!CC Description:  This SUBROUTINE will convert liquid diffusivity from
!CC               units of Pa to units of psi.
!CC
!CC Output Variables:
!CC    LDIFENG =  Liquid Diffusivity (ft2/sec)
!CC
!CC Input Variables:
!CC    LDIFSI =   Liquid Diffusivity (m2/sec)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/16/94)
!CC
!CC********************************************************************

      SUBROUTINE LDIFFCNV(LDIFENG,LDIFSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDIFFCNV
!MS$ ATTRIBUTES ALIAS:'_LDIFFCNV'::LDIFFCNV
!MS$ ATTRIBUTES REFERENCE::LDIFENG,LDIFSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION LDIFENG,LDIFSI

         LDIFENG = LDIFSI * (3.2808D0**2)

      END

!CC********************************************************************


