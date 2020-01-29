!C*********************************************************************
!CC
!CC                       QH2OPT2
!CC
!CC Description:  This subroutine will calculate the water flow rate
!CC               given tower area and liquid mass loading rate.
!CC
!CC Output Variables:
!CC    QW =       Water flow rate (m3/sec)
!CC
!CC Input Variables:
!CC    ML =       Water Mass Loading Rate (kg/m2/sec)
!CC    DL =       Water Density (kg/m3)
!CC    AREA =     Tower Area (m2)
!C*********************************************************************

      SUBROUTINE QH2OPT2(QW,ML,DL,AREA)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::QH2OPT2
!MS$ ATTRIBUTES ALIAS:'_QH2OPT2':: QH2OPT2
!MS$ ATTRIBUTES REFERENCE::QW,ML,DL,AREA

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION QW,ML,DL,AREA

         QW = ML*AREA/DL

      END

!C*********************************************************************

