!C********************************************************************
!C
!C                             LDIFFCNV
!C          CONVERT LIQUID DIFFUSIVITY UNITS FROM m2/sec TO ft2/sec
!C
!C Description:  This SUBROUTINE will convert liquid diffusivity from
!C               units of Pa to units of psi.
!C
!C Output Variables:
!C    LDIFENG =  Liquid Diffusivity (ft2/sec)
!C
!C Input Variables:
!C    LDIFSI =   Liquid Diffusivity (m2/sec)
!C
!C History:
!C    Function written by D. Hokanson (6/16/94)
!C
!C********************************************************************

SUBROUTINE LDIFFCNV(LDIFENG,LDIFSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDIFFCNV
!MS$ ATTRIBUTES ALIAS:'_LDIFFCNV':: LDIFFCNV
!MS$ ATTRIBUTES REFERENCE:: LDIFENG,LDIFSI

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION LDIFENG,LDIFSI

LDIFENG = LDIFSI * (3.2808D0**2)

END SUBROUTINE

!C********************************************************************
