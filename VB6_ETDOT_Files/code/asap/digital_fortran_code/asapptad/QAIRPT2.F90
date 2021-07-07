!C*********************************************************************
!CC
!CC                       QAIRPT2
!CC
!CC Description:  This subroutine will calculate the air flow rate
!CC               given tower area and air mass loading rate.
!CC
!CC Output Variables:
!CC    QA =       Air flow rate (m3/sec)
!CC
!CC Input Variables:
!CC    GM =       Air Mass Loading Rate (kg/m2/sec)
!CC    DG =       Air Density (kg/m3)
!CC    AREA =     Tower Area (m2)
!C*********************************************************************

      SUBROUTINE QAIRPT2(QA,GM,DG,AREA)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::QAIRPT2
!MS$ ATTRIBUTES ALIAS:'_QAIRPT2':: QAIRPT2
!MS$ ATTRIBUTES REFERENCE::QA,GM,DG,AREA

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION QA,GM,DG,AREA

         QA = GM*AREA/DG

      END

!C*********************************************************************

