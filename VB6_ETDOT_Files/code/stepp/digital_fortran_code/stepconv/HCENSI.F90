!CC*******************************************************************
!CC
!CC                               HCENSI
!CC             CONVERT HENRY'S CONSTANT UNITS FROM (-) TO (-)
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for Henry's Constant.  Right now, the units are
!CC               dimensionless in both cases so there is no conversion
!CC               performed.  However, the routine is included in case
!CC               we are manipulating different units in the future.
!CC
!CC Output Variables:
!CC    HCSI =     Henry's Constant (-)
!CC
!CC Input Variables:
!CC    HCENG =    Henry's Constant (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE HCENSI(HCSI,HCENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HCENSI
!MS$ ATTRIBUTES ALIAS:'_HCENSI'::HCENSI
!MS$ ATTRIBUTES REFERENCE::HCSI,HCENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION HCENG, HCSI  
        HCSI = HCENG * 1.0D0                   
      END
 
!CC*******************************************************************


       
