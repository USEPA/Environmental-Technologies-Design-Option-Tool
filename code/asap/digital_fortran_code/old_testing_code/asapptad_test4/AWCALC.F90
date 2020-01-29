!C***************************************************************
!CC
!CC                            AWCALC
!CC    CALCULATION OF THE WETTED SURFACE AREA OF THE PACKING
!CC
!CC Description:  This subroutine will calculate the wetted
!CC               surface area of the packing (AW) in m2/m3.
!CC               The equation to be used is:
!CC
!CC                  AW = AT*(1-exp(-1.45*((STC/ST)^.75)*(RE^.1)*
!CC                                 (FR^(-0.5))*(WE^0.2)))
!CC
!CC               where:
!CC                  AT = Total surface area of packing
!CC                  STC = Critical packing surface tension
!CC                  ST = Surface tension of water
!CC                  RE = ML/(AT*VL)
!CC                  FR = (AT*ML^2)/(DL*DL*9.81)
!CC                  WE = ML^2/(DL*AT*ST)
!CC
!CC Output Variables:
!CC    RE =       Reynold's number (dimensionless)
!CC    FR =       Froude number (dimensionless)
!CC    WE =       Weber number (dimensionless)
!CC    AW =       Wetted surface area of packing (m2/m3)
!CC
!CC Input Variables:
!CC    ML =       Liquid mass loading rate (kg/m^2/sec)
!CC    AT =       Specific surface area of the packing (m^2/m^3)
!CC    VL =       Liquid viscosity (kg/m/sec)
!CC    DL =       Liquid density (kg/m^3)
!CC    ST =       Surface tension of water (kg/sec^2)
!CC    STC =      Critical surface tension of the packing (N/m)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE AWCALC (AW,STC,ST,ML,AT,VL,DL,RE,FR,WE)   
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AWCALC
!MS$ ATTRIBUTES ALIAS:'_AWCALC@40':: AWCALC
!MS$ ATTRIBUTES REFERENCE::AW,STC,ST,ML,AT,VL,DL,RE,FR,WE

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION STC,ST,ML,AT,VL,DL,AW,RE,FR,WE          

         RE = ML/(AT*VL)
         FR = (AT*(ML**2))/(DL*DL*9.81D0)   
         WE = (ML**2)/(DL*AT*ST)                                
         AW = AT*(1.0D0-EXP(-1.45D0*((STC/ST)**0.75)*(RE**0.1)*(FR**(-0.05))*(WE**0.2)))             
      END                                                    
                                                            
!C***************************************************************

