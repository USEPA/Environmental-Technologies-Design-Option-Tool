!C********************************************************************
!C
!C                             ADENSCNV
!C              AIR DENSITY UNITS FROM Kg/m3 TO LBm/Ft3  
!C
!C Description:  This SUBROUTINE will convert Air Density from 
!C               units of Kg/m3 to units of LBm/Ft3.
!C
!C Output Variables:
!C    ADENG =    Air Density (LBm/Ft3)
!C
!C Input Variables:
!C    ADSI =     Air Density (Kg/m3)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE ADENSCNV(ADENG,ADSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ADENSCNV
!MS$ ATTRIBUTES ALIAS:'_ADENSCNV':: ADENSCNV
!MS$ ATTRIBUTES REFERENCE::ADENG,ADSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION ADENG, ADSI

         ADENG = ADSI * 2.20462D0/35.3145D0  

END SUBROUTINE

!C********************************************************************
