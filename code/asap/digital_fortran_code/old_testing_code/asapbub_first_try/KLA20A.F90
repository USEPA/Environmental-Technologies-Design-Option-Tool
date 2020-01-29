!C******************************************************************
!CC
!CC                        KLA20A
!CC
!CC Description:  This subroutine calculates the apparent mass
!CC               transfer coefficient of oxygen at 20 Deg C, KLA20.
!CC
!CC Output Variables:
!CC    V =        Water volume in each tank (L)
!CC    KLA20 =    Apparent oxygen mass transfer coeff. at 20 Deg C (1/sec)
!CC
!CC Input Variables:
!CC    VM3 =      Water volume in each tank (m^3)
!CC    CSTR20 =   DO saturation concentration attained at infinite
!CC               time (mg/L)
!CC    SOTR =     Standardized oxygen mass transfer rate (kg/d)
!CC         =     Rate of oxygen mass transfer at zero D.O. and 20 Deg C
!CC
!CC
!C******************************************************************

      SUBROUTINE KLA20A(KLA20,V,VM3,CSTR20,SOTR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KLA20A
!MS$ ATTRIBUTES ALIAS:'_KLA20A':: KLA20A
!MS$ ATTRIBUTES REFERENCE::KLA20,V,VM3,CSTR20,SOTR

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION KLA20,V,VM3,CSTR20

         V = VM3 * 1000.0D0
         KLA20 = SOTR * 1.0D6 / V / CSTR20 /24.0D0/60.0D0/60.0D0
         
      END

!C******************************************************************

