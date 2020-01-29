!C***************************************************************
!CC
!CC                       LDAIRPT2
!CC
!CC Description:  This subroutine will calculate air mass loading
!CC               rate.
!CC
!CC Output Variables:
!CC    GM =       Air mass loading rate (kg/m^2/sec)
!CC
!CC Input Variables:
!CC    QA =       Air flow rate (m^3/sec)
!CC    DG =       Air density (kg/m^3)
!CC    AREA =     Tower area (m^2)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE LDAIRPT2(GM,QA,DG,AREA)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDAIRPT2
!MS$ ATTRIBUTES ALIAS:'_LDAIRPT2':: LDAIRPT2
!MS$ ATTRIBUTES REFERENCE::GM,QA,DG,AREA

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION GM,QA,DG,AREA

      GM = QA*DG/AREA

      END

!C***************************************************************

