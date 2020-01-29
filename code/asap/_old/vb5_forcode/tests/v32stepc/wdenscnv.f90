!C********************************************************************
!C
!C                             WDENSCNV
!C            WATER DENSITY UNITS FROM Kg/m3 TO LBm/Ft3  
!C
!C Description:  This SUBROUTINE will convert Water Density from 
!C               units of Kg/m3 to units of LBm/Ft3.
!C
!C Output Variables:
!C    WDENG =    Water Density (LBm/Ft3)
!C
!C Input Variables:
!C    WDSI =     Water Density (Kg/m3)
!C
!C History:
!C    Function written by D. Hokanson (6/14/94)
!C
!C********************************************************************

SUBROUTINE WDENSCNV(WDENG,WDSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::WDENSCNV
!MS$ ATTRIBUTES ALIAS:'_WDENSCNV':: WDENSCNV
!MS$ ATTRIBUTES REFERENCE::WDENG,WDSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION WDENG, WDSI

WDENG = WDSI * 2.20462D0/35.3145D0  

END SUBROUTINE

!C********************************************************************
