!C***************************************************************
!CC
!CC                       LDH2OPT2
!CC
!CC Description:  This subroutine will calculate water mass loading
!CC               rate.
!CC
!CC Output Variables:
!CC    ML =       Liquid (water) mass loading rate (kg/m^2/sec)
!CC
!CC Input Variables:
!CC    QW =       Water flow rate (m^3/sec)
!CC    DL =       Water density (kg/m^3)
!CC    AREA =     Tower area (m^2)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE LDH2OPT2(ML,QW,DL,AREA)             
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::LDH2OPT2
!MS$ ATTRIBUTES ALIAS:'_LDH2OPT2':: LDH2OPT2
!MS$ ATTRIBUTES REFERENCE::ML,QW,DL,AREA
    
         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION ML,QW,DL,AREA

         ML = QW*DL/AREA

      END

!C***************************************************************

