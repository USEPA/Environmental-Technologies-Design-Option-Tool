!C********************************************************************
!C
!C                             AQSCONV
!C          CONVERT AQUEOUS SOLUBILITY UNITS FROM (-) TO (-).
!C
!C Description:  This SUBROUTINE will handle the conversion of units
!C               for aqueous solubility.  Right now, the units
!C               are dimensionless in both cases so there is no
!C               conversion performed.  However, the routine is 
!C               included in case we are manipulating different
!C               units in the future.
!C
!C Output Variables:
!C    ASENG =    aqueous solubility (-)
!C
!C Input Variables:
!C    ASSI =     aqueous solubility (-)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE AQSCONV(ASENG,ASSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AQSCONV
!MS$ ATTRIBUTES ALIAS:'_AQSCONV':: AQSCONV
!MS$ ATTRIBUTES REFERENCE::ASENG,ASSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ASENG, ASSI

         ASENG = ASSI * 1.0D0                 

END SUBROUTINE

!C********************************************************************
