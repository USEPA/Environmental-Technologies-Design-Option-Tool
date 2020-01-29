!C****************************************************************
!CC
!CC                    ONDKLAPT
!CC
!CC Description:  This subroutine will calculate the overall
!CC               mass transfer coefficient, KLAOND, using the
!CC               Onda correlation.
!CC
!CC Output Variable:
!CC    KLAOND =   Overall mass transfer coefficient (1/sec) from
!CC               the Onda correlation
!CC    RL =       Liquid phase mass transfer resistance (sec)
!CC    RG =       Gas phase mass transfer resistance (sec)
!CC    RT =       Total mass transfer resistance (sec)
!CC
!CC Input Variables:
!CC    KL =       Liquid phase mass transfer coefficient (m/sec)
!CC    AW =       Wetted surface area of packing (m2/m3)
!CC    KG =       Gas phase mass transfer coefficient (m/sec)
!CC    HC =       Henry's constant (dimensionless)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C****************************************************************

      SUBROUTINE ONDKLAPT(KLAOND,RL,RG,RT,KL,AW,KG,HC)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::ONDKLAPT
!MS$ ATTRIBUTES ALIAS:'_ONDKLAPT@32':: ONDKLAPT
!MS$ ATTRIBUTES REFERENCE::KLAOND,RL,RG,RT,KL,AW,KG,HC

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION KLAOND,RL,RG,RT,KL,AW,KG,HC 

         RL = 1/(KL*AW)                                        
         RG = 1/(KG*AW*HC)                                    
         RT = RL+RG                                          
         KLAOND = 1/(RL+RG)                                    

      END

!C****************************************************************

