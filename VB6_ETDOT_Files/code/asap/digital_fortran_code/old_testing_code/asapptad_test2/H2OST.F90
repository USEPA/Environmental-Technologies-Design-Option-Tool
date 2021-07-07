!C*****************************************************************
!CC
!CC                          H2OST
!CC
!CC Description:  This subroutine will calculate the surface tension
!CC               of water, given temperature
!CC
!CC Output Variable:
!CC    ST =       Surface tension of water (kg/sec^2)
!CC
!CC Input Variable:
!CC    TEMP =     Operating temperature (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C*****************************************************************

      SUBROUTINE H2OST(ST,TEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::H2OST
!MS$ ATTRIBUTES ALIAS:'_H2OST':: H2OST
!MS$ ATTRIBUTES REFERENCE::ST,TEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION ST,TEMP

         ST = (7.558301D-2) - (1.3143D-4)*(TEMP-273) - (4.7616D-7)*((TEMP-273)**2)

      END

!C*****************************************************************
                      
