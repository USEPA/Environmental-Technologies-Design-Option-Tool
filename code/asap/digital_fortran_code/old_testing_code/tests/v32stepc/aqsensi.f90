!C*******************************************************************
!C
!C                               AQSENSI 
!C             CONVERT AQUEOUS SOLUBILITY UNITS FROM (-) TO (-)  
!C
!C Description:  This SUBROUTINE will handle the conversion of units        
!C               for aqueous solubility.  Right now, the units are
!C               dimensionless in both cases so there is no conversion
!C               performed.  However, the routine is included in case 
!C               we are manipulating different units in the future.                                  
!C
!C Output Variables:
!C    AQSSI =     Aqueous Solubility (-)
!C
!C Input Variables:
!C    AQSENG =    Aqueous Solubility (-)
!C
!C History:
!C    Function written by D. Hokanson (6/23/94)
!C
!C*******************************************************************

SUBROUTINE AQSENSI(AQSSI,AQSENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AQSENSI
!MS$ ATTRIBUTES ALIAS:'_AQSENSI':: AQSENSI
!MS$ ATTRIBUTES REFERENCE::AQSSI,AQSENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION AQSENG, AQSSI  
        AQSSI = AQSENG * 1.0D0                   

END SUBROUTINE

!C*******************************************************************


       
