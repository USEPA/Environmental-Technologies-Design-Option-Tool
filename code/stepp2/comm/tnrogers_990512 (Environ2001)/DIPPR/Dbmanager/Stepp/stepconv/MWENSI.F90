!CC*******************************************************************
!CC
!CC                               MWENSI
!CC             CONVERT MOLECULAR WEIGHT UNITS FROM (-) TO (-)
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for molecular weight.  Right now, the units are
!CC               dimensionless in both cases so there is no conversion
!CC               performed.  However, the routine is included in case
!CC               we are manipulating different units in the future.
!CC
!CC Output Variables:
!CC    MWSI =     Molecular Weight (-)
!CC
!CC Input Variables:
!CC    MWENG =    Molecular Weight (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE MWENSI(MWSI,MWENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::MWENSI
!MS$ ATTRIBUTES ALIAS:'_MWENSI'::MWENSI
!MS$ ATTRIBUTES REFERENCE::MWSI,MWENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION MWENG, MWSI  
        MWSI = MWENG * 1.0D0                   
      END
 
!CC*******************************************************************


       
