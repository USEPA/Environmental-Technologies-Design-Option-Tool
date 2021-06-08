!C*******************************************************************
!C
!C                                LDIFENSI
!C          CONVERT LIQUID DIFFUSIVITY FROM Ft2/sec to m2/sec
!C                    
!C
!C Description:  This SUBROUTINE will convert liquid diffusivity             
!C               from units of Ft2/sec to m2/sec
!C
!C Output Variables:
!C    LDSI =     Liquid Diffusivity (m2/sec)                 
!C
!C Input Variables:
!C    LDENG =    Liquid Diffusivity (Ft2/sec)                   
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE LDIFENSI(LDSI,LDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDIFENSI
!MS$ ATTRIBUTES ALIAS:'_LDIFENSI':: LDIFENSI
!MS$ ATTRIBUTES REFERENCE:: LDSI,LDENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION LDENG, LDSI

LDSI = LDENG / (3.2808D0**2)                       

END SUBROUTINE

!C*******************************************************************



