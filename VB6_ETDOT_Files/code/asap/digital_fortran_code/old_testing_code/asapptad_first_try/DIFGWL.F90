!C**************************************************************
!CC
!CC                   DIFGWL
!CC
!CC Description:  This subroutine will calculate the gas diffusivity
!CC               using the Wilke-Lee modification of the Hirschfelder-
!CC               Bird-Spotz method.
!CC
!CC Output Variable:
!CC    DIFG =     Gas Diffusivity (m^2/sec)
!CC
!CC Input Variables:
!CC    MW =       Molecular weight of compound
!CC    VB =       Molar volume of compound (m^3/kmol)
!CC    TEMPB =    Boiling point temperature of compound (K)
!CC    TEMP =     Operating temperature (K)
!CC    PRES =     Operating pressure (atm)
!CC
!CC Variables internal to Subroutine DIFGWL:
!CC    MA =       Molecular weight of air
!CC    MB =       Molecular weight of compound
!CC    RA =       Molecular radius of air
!CC    RB =       Molecular radius of compound
!CC    RAB =      Molecular separation at collision (nm)
!CC    EKB =      Eb/k
!CC    EKA =      Ea/k
!CC    EKEAB =    Energy of molecular attraction / Boltzmann's constant
!CC    TKEAB =    kT/Eab
!CC    EE =       Temporary value = log(TKEAB)
!CC    YVAL =     f(kT/Eab) = Collision function
!CC    PT =       Absolute pressure (N/m^2)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C**************************************************************

      SUBROUTINE DIFGWL(DIFG,MW,VB,TEMPB,TEMP,PRES)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DIFGWL
!MS$ ATTRIBUTES ALIAS:'_DIFGWL':: DIFGWL
!MS$ ATTRIBUTES REFERENCE::DIFG,MW,VB,TEMPB,TEMP,PRES

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFG,MW,VB,TEMPB,TEMP,PRES,MB,MA,RB,RA,RAB,EKB,EKA,EKEAB,TKEAB,EE,YVAL,SQM,PT 
     
         MB = MW
         MA = 28.95D0                                            
         RB = 1.18D0*((VB)**0.33333)                          
         RA = 0.3711D0                                       
         RAB = (RA+RB)/2.0D0                                  
         EKB = 1.21D0*TEMPB
         EKA = 78.6D0                                           
         EKEAB = (EKB*EKA)**0.5                                
         TKEAB = TEMP/EKEAB
         EE = LOG(TKEAB)/2.303D0
         YVAL = 10**(-0.14329D0 - 0.48343D0*(EE) + 0.1939D0*(EE)**2+ 0.13612D0*(EE)**3 - 0.20578D0*(EE)**4 + 0.083899D0*(EE)**5 - 0.011491D0*(EE)**6)
         SQM = (1/MA+1/MB)**0.5 
         PT = 101325.0D0*PRES  
         DIFG = (0.0001D0*(1.084D0-(0.249D0*SQM))*(TEMP**1.5)*SQM)/(PT*RAB*YVAL*RAB)

      END

!C**************************************************************

