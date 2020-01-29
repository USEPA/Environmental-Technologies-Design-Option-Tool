!CC********************************************************************
!CC
!CC                             TEMPCNV
!CC              CONVERT TEMPERATURE UNITS FROM C TO F
!CC
!CC Description:  This SUBROUTINE will convert a temperature
!CC               from units of C to units of F.
!CC
!CC Output Variables:
!CC    TEMPENG =  Temperature (F)
!CC
!CC Input Variables:
!CC    TEMPSI =   Temperature (C)
!CC
!CC History:
!CC    Function written by D. Hokanson (6/14/94)
!CC
!CC********************************************************************

      SUBROUTINE TEMPCNV(TEMPENG,TEMPSI)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::TEMPCNV
!MS$ ATTRIBUTES ALIAS:'_TEMPCNV'::TEMPCNV
!MS$ ATTRIBUTES REFERENCE::TEMPENG,TEMPSI

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION TEMPENG, TEMPSI

         TEMPENG = TEMPSI * (9.0D0/5.0D0) + 32.0D0         

      END

!CC********************************************************************


