!C********************************************************************
!C
!C                             RICONV
!C          CONVERT REFRACTIVE INDEX UNITS FROM (-) TO (-).
!C
!C Description:  This SUBROUTINE will handle the conversion of units
!C               for refractive index.  Right now, the units
!C               are dimensionless in both cases so there is no
!C               conversion performed.  However, the routine is 
!C               included in case we are manipulating different
!C               units in the future.
!C
!C Output Variables:
!C    RIENG =    refractive index (-)
!C
!C Input Variables:
!C    RISI =     refractive index (-)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE RICONV(RIENG,RISI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::RICONV
!MS$ ATTRIBUTES ALIAS:'_RICONV':: RICONV
!MS$ ATTRIBUTES REFERENCE::RIENG,RISI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION RIENG, RISI

RIENG = RISI * 1.0D0                 

END SUBROUTINE

!C********************************************************************
