!C*******************************************************************
!C
!C                               RIENSI 
!C             CONVERT REFRACTIVE INDEX UNITS FROM (-) TO (-)  
!C
!C Description:  This SUBROUTINE will handle the conversion of units        
!C               for refractive index.  Right now, the units are
!C               dimensionless in both cases so there is no conversion
!C               performed.  However, the routine is included in case 
!C               we are manipulating different units in the future.                                  
!C
!C Output Variables:
!C    RISI =     Refractive Index (-)
!C
!C Input Variables:
!C    RIENG =    Refractive Index (-)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE RIENSI(RISI,RIENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::RIENSI
!MS$ ATTRIBUTES ALIAS:'_RIENSI':: RIENSI
!MS$ ATTRIBUTES REFERENCE::RISI,RIENG


      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION RIENG, RISI  

RISI = RIENG * 1.0D0                   
END	SUBROUTINE
 
!C*******************************************************************


       
