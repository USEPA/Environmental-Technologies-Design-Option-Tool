!C********************************************************************
!CC
!CC                        TRUEKLA
!CC
!CC Description:  This subroutine will return values for the true
!CC               oxygen mass transfer coefficient at the operating
!CC               temperature (KLATT), for the true oxygen mass
!CC               transfer coefficient at 20 deg C (KLAT20, and for
!CC               PHI (a value used in calculating the true oxygen
!CC               mass transfer coefficient).
!CC
!CC Output Variables:
!CC    KLAO2 =    True oxygen mass transfer coefficient (1/sec)
!CC               at operating temperature, TEMP
!CC    KLAT20 =   True oxygen mass transfer coeff. at 20 Deg C (1/sec)
!CC    PHI =      Parameter used in calculating true Kla (1/sec)
!CC
!CC Input Variables:
!CC    QAIR =     Air flow rate (std m^3/hr) --> 20 Deg C, 1 atm, 36% r.h.
!CC    V =        Water Volume in each tank (L)
!CC    PB =       Barometric pressure (atm)
!CC    GAMMAW =   Weight density of water
!CC    KLA20 =    Apparent oxygen mass transfer coeff. at 20 Deg C (1/min)
!CC    DEFF =     Effective saturation depth (m)
!CC    TEMP =     Operating temperature (Deg K)
!CC
!CC Variables Internal to Subroutine TRUEKLA
!CC    MO =       Molecular weight of oxygen (O2)
!CC    MA =       Molecular weight of air
!CC    RHOA =     Density of air (mg/L)
!CC    QAIRLM =   Volumetric air flow rate (L/min)
!CC    HO =       Henry's constant for O2
!CC    TEMPC =    Temperature in Deg C
!CC    THETA =    Argument for Van't Hoff temperature relationship
!CC    KLATT =    True Kla at operating temperature, TEMP
!CC
!C********************************************************************

      SUBROUTINE TRUEKLA(KLAO2,KLAT20,PHI,QAIR,V,PB,GAMMAW,KLA20,DEFF,TEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::TRUEKLA
!MS$ ATTRIBUTES ALIAS:'_TRUEKLA':: TRUEKLA
!MS$ ATTRIBUTES REFERENCE::KLAO2,KLAT20,PHI,QAIR,V,PB,GAMMAW,KLA20,DEFF,TEMP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION KLAO2,KLAT20,PHI,QAIR,V,PB,GAMMAW,KLA20,DEFF,TEMP,MO,MA,RHOA,QAIRLM,HO,TEMPC,THETA,KLATT

            MO = 32.0D0
            MA = 28.87D0
            RHOA = 1240.0D0
            QAIRLM = QAIR * 1000.0D0 / 60.0D0
            HO = 50.0D0
            PHI = (MO*RHOA*QAIRLM)/(MA*HO*V*(PB+GAMMAW/144.0D0/14.696D0*DEFF*3.2808D0))/60.0D0
            KLAT20 = KLA20/(1-KLA20/(2.0D0*PHI))
            TEMPC = TEMP - 273.15D0
            THETA = 1.024D0
            KLATT = KLAT20 * THETA**(TEMPC-20.0D0)
            KLAO2 = KLATT

      END
            
!C********************************************************************

