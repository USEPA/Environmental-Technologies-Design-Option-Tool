!C*******************************************************************
!C
!C                                GDIFENSI
!C          CONVERT GAS DIFFUSIVITY FROM Ft2/sec to m2/sec
!C                    
!C
!C Description:  This SUBROUTINE will convert gas diffusivity             
!C               from units of Ft2/sec to m2/sec
!C
!C Output Variables:
!C    GDSI =     Gas Diffusivity (m2/sec)                 
!C
!C Input Variables:
!C    GDENG =    Gas Diffusivity (Ft2/sec)                   
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE GDIFENSI(GDSI,GDENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GDIFENSI
!MS$ ATTRIBUTES ALIAS:'_GDIFENSI':: GDIFENSI
!MS$ ATTRIBUTES REFERENCE::GDSI,GDENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION GDENG, GDSI

GDSI = GDENG / (3.2808D0**2)                       

END SUBROUTINE
 
!C*******************************************************************


       
