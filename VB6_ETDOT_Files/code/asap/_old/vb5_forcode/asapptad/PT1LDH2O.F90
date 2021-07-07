!C***************************************************************
!CC
!CC                       PT1LDH2O
!CC
!CC Description:  This subroutine will calculate water mass loading
!CC               rate.
!CC
!CC Output Variables:
!CC    ML =       Liquid (water) mass loading rate (kg/m^2/sec)
!CC
!CC Input Variables:
!CC    VQ =       Air to water ratio (dimensionless)
!CC    DG =       Air density (kg/m^3)
!CC    DL =       Water density (kg/m^3)
!CC    GM =       Air mass loading rate (kg/m^2/sec)
!CC
!CC Variable internal to Subroutine LDH2OOT1:
!CC    VQM =      Air to water ratio on a mass basis
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PT1LDH2O(ML,VQ,DG,DL,GM)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PT1LDH2O
!MS$ ATTRIBUTES ALIAS:'_PT1LDH2O@20':: PT1LDH2O
!MS$ ATTRIBUTES REFERENCE::ML,VQ,DG,DL,GM
    
      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION VQM,ML,VQ,DG,DL,GM

         VQM = VQ*(DG/DL)                            
         ML = GM/VQM                                

      END

!C***************************************************************
 
