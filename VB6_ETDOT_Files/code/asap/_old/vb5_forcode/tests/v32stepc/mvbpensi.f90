!C*******************************************************************
!C
!C                                MVBPENSI
!C          CONVERT MOLAR VOLUME AT NORMAL BOILING POINT FROM LBm/Ft3
!C               Kg/m3
!C
!C Description:  This SUBROUTINE will convert molar volume at normal         
!C               boiling point from units of LBm/Ft3 to Kg/m3.
!C
!C Output Variables:
!C    MVBSI =     Molar Volume at Normal Boiling Point (Kg/m3)
!C
!C Input Variables:
!C    MVBENG =    Molar Volume at Normal Boiling Point (LBm/Ft3)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE MVBPENSI(MVBSI,MVBENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MVBPENSI
!MS$ ATTRIBUTES ALIAS:'_MVBPENSI':: MVBPENSI
!MS$ ATTRIBUTES REFERENCE:: MVBSI,MVBENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION MVBENG, MVBSI

MVBSI = MVBENG * 2.20462D0 / 35.3145D0              

END SUBROUTINE

!C*******************************************************************


       
