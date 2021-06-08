!CC*******************************************************************
!CC
!CC                               RIENSI
!CC             CONVERT REFRACTIVE INDEX UNITS FROM (-) TO (-)
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for refractive index.  Right now, the units are
!CC               dimensionless in both cases so there is no conversion
!CC               performed.  However, the routine is included in case
!CC               we are manipulating different units in the future.
!CC
!CC Output Variables:
!CC    RISI =     Refractive Index (-)
!CC
!CC Input Variables:
!CC    RIENG =    Refractive Index (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE RIENSI(RISI,RIENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::RIENSI
!MS$ ATTRIBUTES ALIAS:'_RIENSI'::RIENSI
!MS$ ATTRIBUTES REFERENCE::RISI,RIENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION RIENG, RISI  
        RISI = RIENG * 1.0D0                   
      END
 
!CC*******************************************************************


       
