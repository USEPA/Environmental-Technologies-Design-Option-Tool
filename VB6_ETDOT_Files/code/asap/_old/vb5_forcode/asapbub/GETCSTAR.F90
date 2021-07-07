!C*********************************************************************
!CC
!CC                   GETCSTAR
!CC
!CC Description:  This subroutine will calculate a value of CSTR20.
!CC
!CC Output Variable:
!CC    CSTR20 =   DO saturation concentration attained at infinite
!CC               time (mg/L)
!CC    GAMMAW =   Weight density of water
!CC    DEFF =     Effective saturation depth (m)
!CC
!CC Input Variables:
!CC    PB =       Barometric pressure (atm)
!CC    DEPTHW =   Water depth (m)
!CC
!CC Variables Internal to Subroutine GETCSTR
!CC    CSTRS =    Tabular value of D.O. surface saturation conc. at 20 C
!CC    PV =       Vapor pressure of water (atm)
!CC    PS =       Standard barometric pressure of 1.00 atm
!CC
!C********************************************************************

      SUBROUTINE GETCSTAR(CSTR20,GAMMAW,DEFF,PB,DEPTHW)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::GETCSTAR
!MS$ ATTRIBUTES ALIAS:'_GETCSTAR':: GETCSTAR
!MS$ ATTRIBUTES REFERENCE::CSTR20,GAMMAW,DEFF,PB,DEPTHW

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION CSTR20,PB,DEPTHW,CSTRS,PV,GAMMAW,DEFF,PS

         PV = 0.023D0
         GAMMAW = 62.4D0
         DEFF = DEPTHW / 3.0D0       
         PS = 1.0D0
         CSTRS = 9.09D0 
         CSTR20 = CSTRS * ((PB-PV+(GAMMAW/144.0D0/14.696D0)*DEFF*3.2808D0)/(PS-PV))

      END

!C********************************************************************

