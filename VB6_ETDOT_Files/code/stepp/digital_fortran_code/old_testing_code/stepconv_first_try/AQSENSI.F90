!CC*******************************************************************
!CC
!CC                               AQSENSI
!CC             CONVERT AQUEOUS SOLUBILITY UNITS FROM (-) TO (-)
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for aqueous solubility.  Right now, the units are
!CC               dimensionless in both cases so there is no conversion
!CC               performed.  However, the routine is included in case
!CC               we are manipulating different units in the future.
!CC
!CC Output Variables:
!CC    AQSSI =     Aqueous Solubility (-)
!CC
!CC Input Variables:
!CC    AQSENG =    Aqueous Solubility (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE AQSENSI(AQSSI,AQSENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AQSENSI
!MS$ ATTRIBUTES ALIAS:'_AQSENSI'::AQSENSI
!MS$ ATTRIBUTES REFERENCE::AQSSI,AQSENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION AQSENG, AQSSI  
        AQSSI = AQSENG * 1.0D0                   
      END
 
!CC*******************************************************************


       
