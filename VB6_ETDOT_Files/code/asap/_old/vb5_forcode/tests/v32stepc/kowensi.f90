!C*******************************************************************
!C
!C                               KOWENSI 
!C             CONVERT OCTANOL WATER PARTITION COEFF FROM (-) TO (-)  
!C
!C Description:  This SUBROUTINE will handle the conversion of units        
!C               for octanol water partition coeff.  Right now, the 
!C               units are dimensionless in both cases so there is no           
!C               conversion performed.  However, the routine is included      
!C               incase we are manipulating different units in the future                                  
!C
!C Output Variables:
!C    KOWSI =     Octanol Water Partition Coeff (-)
!C
!C Input Variables:
!C    KOWENG =    Octanol Water Partition Coeff (-)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE KOWENSI(KOWSI,KOWENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KOWENSI
!MS$ ATTRIBUTES ALIAS:'_KOWENSI':: KOWENSI
!MS$ ATTRIBUTES REFERENCE::KOWSI,KOWENG

	IMPLICIT DOUBLE PRECISION (A-H,O-Z)
	DOUBLE PRECISION KOWENG, KOWSI  

KOWSI = KOWENG * 1.0D0                   

END SUBROUTINE

!C*******************************************************************


       
