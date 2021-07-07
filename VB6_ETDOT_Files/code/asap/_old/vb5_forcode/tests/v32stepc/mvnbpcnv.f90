!C********************************************************************
!C
!C                             MVNBPCNV
!C          CONVERT MOLAR VOLUME AT NORMAL BOILING TEMP FROM M3/KMOL TO
!C          FT3/LBm-MOL
!C
!C Description:  This SUBROUTINE will convert molar volume at 
!C               normal boiling point units from m3/Kmol to Ft3/LBm-mol.
!C
!C Output Variables:
!C    MVNENG =    Molar Volume at Normal Boiling Pt (Ft3/LBm-mol)
!C
!C Input Variables:
!C    MVNSI =     Molar Volume at Normal Boiling Pt (m3/Kmol)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE MVNBPCNV(MVNENG,MVNSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVNBPCNV
!MS$ ATTRIBUTES ALIAS:'_MVNBPCNV':: MVNBPCNV
!MS$ ATTRIBUTES REFERENCE::MVNENG,MVNSI

		IMPLICIT DOUBLE PRECISION (A-H,O-Z)
		DOUBLE PRECISION MVNENG, MVNSI

 MVNENG = MVNSI * 35.3145D0/2.20462D0           

END SUBROUTINE

!C********************************************************************
