!C*******************************************************************
!C
!C                                MVOTENSI
!C               CONVERT MOLAR VOLUME AT OPERATING TEMP FROM LBm/Ft3
!C               Kg/m3
!C
!C Description:  This SUBROUTINE will convert molar volume at operating
!C               temp from units of LBm/Ft3 to Kg/m3
!C
!C Output Variables:
!C    MVOSI =     Molar Volume at Operating Temp (Kg/m3)
!C
!C Input Variables:
!C    MVOENG =    Molar Volume at Operating Temp (LBm/Ft3)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE MVOTENSI(MVOSI,MVOENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVOTENSI
!MS$ ATTRIBUTES ALIAS:'_MVOTENSI':: MVOTENSI
!MS$ ATTRIBUTES REFERENCE::MVOSI,MVOENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION MVOENG, MVOSI

MVOSI = MVOENG * 2.20462D0 / 35.3145D0              

END SUBROUTINE

!C*******************************************************************



