!C**************************************************************
!CC
!CC                   DIFGWL
!CC
!CC Description:  This subroutine will calculate the gas
!CC               diffusivity using the Wilke-Lee modification
!CC               of the Hirschfelder-Bird-Spotz method.
!CC
!CC Output Variables:
!CC    DIFG =     Gas Diffusivity value (m^2/sec)
!CC    ERRORF =   Error flag
!CC    SRCSHT =   Source of this value (Short Version)
!CC    SRCLNG =   Source of this value (Long Version)
!CC    DIFGT =    Temperature of this calculation (C)
!CC
!CC Input Variables:
!CC    MW =       Molecular weight of compound
!CC    VB =       Molar volume of compound (m^3/kmol) at
!CC               normal boiling point
!CC    TNBP =     Boiling point temperature of compound (C)
!CC    TEMPOP =   Temperature of the calculation (C)
!CC    PRES =     Operating pressure (N/m2)
!CC
!CC Variables internal to Subroutine DIFGWL:
!CC    MA =       Molecular weight of air
!CC    MB =       Molecular weight of compound
!CC    RA =       Molecular radius of air
!CC    RB =       Molecular radius of compound
!CC    RAB =      Molecular separation at collision (nm)
!CC    EKB =      Eb/k
!CC    EKA =      Ea/k
!CC    EKEAB =    Energy of molecular attraction / Boltzmann's
!CC               constant
!CC    TKEAB =    kT/Eab
!CC    EE =       Temporary value = log(TKEAB)
!CC    YVAL =     f(kT/Eab) = Collision function
!CC    TEMP =     Temeprature of the calculation (K)
!CC    TEMPB =    Boiling point temp. of compound (K)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC           Modified by D. Hokanson (4/5/94)
!CC
!C**************************************************************

      SUBROUTINE DIFGWL(DIFG,MW,VB,TNBP,TEMPOP,PRES,ERRORF,SRCSHT,SRCLNG,DIFGT)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DIFGWL
!MS$ ATTRIBUTES ALIAS:'_DIFGWL':: DIFGWL
!MS$ ATTRIBUTES REFERENCE::DIFG,MW,VB,TNBP,TEMPOP,PRES,ERRORF,SRCSHT,SRCLNG,DIFGT

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFG,MW,VB,TEMPB,TEMP,PRES,MB,MA,RB
      DOUBLE PRECISION RA,RAB,EKB,EKA,EKEAB,TKEAB,EE,YVAL,SQM
      DOUBLE PRECISION PT,TNBP,TEMPOP
      INTEGER ERRORF,SRCSHT,SRCLNG

         TEMP = TEMPOP + 273.15D0
         TEMPB = TNBP + 273.15D0
         ERRORF = 0
         SRCSHT = 13
         DIFGT = TEMPOP
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
         YVAL = 10**(-0.14329D0-0.48343D0*(EE)+0.1939D0*(EE)**2 + 0.13612D0*(EE)**3 - 0.20578D0*(EE)**4 + 0.083899D0*(EE)**5 - 0.011491D0*(EE)**6)
         SQM = (1/MA+1/MB)**0.5  
         DIFG = (0.0001D0*(1.084D0-(0.249D0*SQM))*(TEMP**1.5)*SQM)/(PRES*RAB*YVAL*RAB)

      END

!C**************************************************************


