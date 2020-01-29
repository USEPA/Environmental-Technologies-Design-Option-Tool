!CC*******************************************************************
!CC
!CC                               KOWENSI
!CC             CONVERT OCTANOL WATER PARTITION COEFF FROM (-) TO (-)
!CC
!CC Description:  This SUBROUTINE will handle the conversion of units
!CC               for octanol water partition coeff.  Right now, the
!CC               units are dimensionless in both cases so there is no
!CC               conversion performed.  However, the routine is included
!CC               incase we are manipulating different units in the future
!CC
!CC Output Variables:
!CC    KOWSI =     Octanol Water Partition Coeff (-)
!CC
!CC Input Variables:
!CC    KOWENG =    Octanol Water Partition Coeff (-)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/23/94)
!CC
!CC*******************************************************************

      SUBROUTINE KOWENSI(KOWSI,KOWENG)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KOWENSI
!MS$ ATTRIBUTES ALIAS:'_KOWENSI'::KOWENSI
!MS$ ATTRIBUTES REFERENCE::KOWSI,KOWENG

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION KOWENG, KOWSI  
        KOWSI = KOWENG * 1.0D0                   
      END
 
!CC*******************************************************************


       
