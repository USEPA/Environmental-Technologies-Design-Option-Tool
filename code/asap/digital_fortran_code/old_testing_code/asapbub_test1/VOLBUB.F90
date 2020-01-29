!C***************************************************************
!CC
!CC                         VOLBUB
!CC
!CC Description:  This subroutine will calculate the tank volume
!CC               required to meet a treatment objective for a
!CC               given chemical.
!CC
!CC Output Variables:
!CC    VOLTNK =   Tank Volume to Meet Treatment Objective (m3)
!CC    ERRORF =   Error Flag (Value of 0 means no error, Value
!CC               of -1 means a negative log would have been taken
!CC
!CC Input Variables:
!CC    HENRYC =   Henry's constant of compound (dimensionless)
!CC    QAIR =     Air flow rate (m3/sec)
!CC    KLA =      Mass Transfer Coefficient of Compound (1/sec)
!CC    CINFL =    Influent Concentration of Compound (ug/L)
!CC    CTO =      Treatment Objective of Compound (ug/L)
!CC    NTANK =    Number of Tanks in Series (-)
!CC    QW =       Water Flow Rate (m3/sec)
!CC
!C***************************************************************

      SUBROUTINE VOLBUB(VOLTNK,HENRYC,QAIR,KLA,CINFL,CTO,NTANK,QW,ERRORF)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VOLBUB
!MS$ ATTRIBUTES ALIAS:'_VOLBUB':: VOLBUB
!MS$ ATTRIBUTES REFERENCE::VOLTNK,HENRYC,QAIR,KLA,CINFL,CTO,NTANK,QW,ERRORF
         
         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         INTEGER NTANK,ERRORF
         DOUBLE PRECISION VOLTNK,HENRYC,QAIR,KLA,CINFL,CTO,QW
         DOUBLE PRECISION PARAM1, PARAM2, PARAM3

         ERRORF = 0
         PARAM1 = (CINFL/CTO)**(1.0D0/DBLE(NTANK))
         PARAM2 = (PARAM1-1.0D0)*QW/QAIR/HENRYC
         PARAM3 = 1.0D0 - PARAM2
         IF (PARAM3.LE.(0.0D0)) THEN
            ERRORF = -1
         ELSE
            VOLTNK = - (HENRYC * QAIR / KLA) * DLOG(PARAM3)
         END IF

      END

!C***************************************************************

