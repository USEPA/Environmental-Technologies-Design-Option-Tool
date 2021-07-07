!C***************************************************************
!CC
!CC                            DIFFL
!CC         CALCULATION OF THE LIQUID DIFFUSIVITY, DIFL
!CC
!CC Description:  This subroutine will determine a value for
!CC               liquid diffusivity, DIFL.  If a value is
!CC               available, it will simply be entered.
!CC               Otherwise, DIFL will be calculated by the
!CC               program.
!CC
!CC Variables:
!CC    MW =       Molecular weight of the compound
!CC    VB =       Molal volume of the compound (m3/kg-mol)
!CC    DIFL =     Liquid Diffusivity (m2/sec)
!CC
!CC
!CC History:
!CC    6/7/93 - changed the REAL data types to DOUBLE PRECISION to give
!CC    8 bytes of precision. (See history note in STEP2.FOR, same date).
!CC       -ry
!CC
!CC    9/30/93 - Modified by DAVID R. HOKANSON to do the diffusivity
!CC              calculations in separate FORTRAN subroutines called
!CC              DIFLPOL.FOR and DIFLHL.FOR
!CC
!CC***************************************************************

      SUBROUTINE DIFFL (DIFL,VL, MW, VB)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DIFFL
!MS$ ATTRIBUTES ALIAS:'_DIFFL':: DIFFL
!MS$ ATTRIBUTES REFERENCE::DIFL,VL,MW,VB

         DOUBLE PRECISION DIFL, MW, VL, VB

            IF (MW.GE.1000) THEN
               CALL DIFLPOL(DIFL,MW)   
            ELSE
               CALL DIFLHL(DIFL,VL,VB)
            END IF

      END
     
!C***************************************************************

