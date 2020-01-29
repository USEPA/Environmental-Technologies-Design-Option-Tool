!C********************************************************************
!C
!C                             MVOPTCNV
!C          CONVERT MOLAR VOLUME AT OPERATING TEMP FROM M3/KMOL TO
!C          FT3/LBm-MOL
!C
!C Description:  This SUBROUTINE will convert molar volume at 
!C               operating temp units from m3/Kmol to Ft3/LBm-mol.
!C
!C Output Variables:
!C    MVOENG =    Molar Volume at Operating Temp (Ft3/LBm-mol)
!C
!C Input Variables:
!C    MVOSI =     Molar Volume at Operating Temp (m3/Kmol)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE MVOPTCNV(MVOENG,MVOSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVOPTCNV
!MS$ ATTRIBUTES ALIAS:'_MVOPTCNV':: MVOPTCNV
!MS$ ATTRIBUTES REFERENCE::MVOENG,MVOSI

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION MVOENG, MVOSI

 MVOENG = MVOSI * 35.3145D0/2.20462D0           

END	SUBROUTINE

!C********************************************************************
