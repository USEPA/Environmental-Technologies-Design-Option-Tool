!CC*******************************************************************
!CC
!CC                               ACENSI
!CC             CONVERT ACTIVITY COEFFICIENT UNITS FROM (-) TO (-)
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for activity coefficient.  Right now, the units are
!CC               dimensionless in both cases so there is no conversion
!CC               performed.  However, the routine is included in case
!CC               we are manipulating different units in the future.
!CC
!CC Output Variables:
!CC    ACSI =     Activity Coefficient (-)
!CC
!CC Input Variables:
!CC    ACENG =    Activity Coefficient (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE ACENSI(ACSI,ACENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ACENSI
!MS$ ATTRIBUTES ALIAS:'_ACENSI'::ACENSI
!MS$ ATTRIBUTES REFERENCE::ACSI,ACENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ACENG, ACSI  
        ACSI = ACENG * 1.0D0                   
      END
 
!CC*******************************************************************


       
