!C********************************************************************
!C
!C                             GDIFFCNV
!C          CONVERT GAS DIFFUSIVITY UNITS FROM m2/sec TO ft2/sec
!C
!C Description:  This SUBROUTINE will convert gas diffusivity from
!C               units of Pa to units of psi.
!C
!C Output Variables:
!C    GDIFENG =  Gas Diffusivity (ft2/sec)
!C
!C Input Variables:
!C    GDIFSI =   Gas Diffusivity (m2/sec)
!C
!C History:
!C    Function written by D. Hokanson (6/16/94)
!C
!C********************************************************************

SUBROUTINE GDIFFCNV(GDIFENG,GDIFSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GDIFFCNV
!MS$ ATTRIBUTES ALIAS:'_GDIFFCNV':: GDIFFCNV
!MS$ ATTRIBUTES REFERENCE::GDIFENG,GDIFSI

		IMPLICIT DOUBLE PRECISION (A-H,O-Z)
		DOUBLE PRECISION GDIFENG,GDIFSI

GDIFENG = GDIFSI * (3.2808D0**2)

END SUBROUTINE

!C********************************************************************
