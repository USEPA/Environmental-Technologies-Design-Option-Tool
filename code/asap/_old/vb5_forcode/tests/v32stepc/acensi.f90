!C*******************************************************************
!C
!C                               ACENSI 
!C             CONVERT ACTIVITY COEFFICIENT UNITS FROM (-) TO (-)  
!C
!C Description:  This SUBROUTINE will handle the conversion of units        
!C               for activity coefficient.  Right now, the units are
!C               dimensionless in both cases so there is no conversion
!C               performed.  However, the routine is included in case 
!C               we are manipulating different units in the future.                                  
!C
!C Output Variables:
!C    ACSI =     Activity Coefficient (-)
!C
!C Input Variables:
!C    ACENG =    Activity Coefficient (-)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE ACENSI(ACSI,ACENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ACENSI
!MS$ ATTRIBUTES ALIAS:'_ACENSI':: ACENSI
!MS$ ATTRIBUTES REFERENCE::ACSI,ACENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ACENG, ACSI  
        ACSI = ACENG * 1.0D0                   

END	SUBROUTINE
 
!C*******************************************************************


       
