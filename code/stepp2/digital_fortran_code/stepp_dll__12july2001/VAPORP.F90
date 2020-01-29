!CC****************************************************************
!CC
!CC                       VAPORP
!CC               DETERMINE VAPOR PRESSURE
!CC
!CC Description:  This subroutine will determine the vapor pressure
!CC               for the compound of interest at the temperature
!CC               of interest.
!CC
!CC Output Variables:
!CC    PVAP =     Vapor pressure of compound at temperature TT
!CC                  (Units:  N/m2 if source is DIPPR801
!CC                           mm Hg if source is Yaws)
!CC
!CC Input Variables:
!CC    TT =       Temperature of interest (K)
!CC    ANTA
!CC    ANTB
!CC    ANTC =      Coefficients for the vapor pressure correlation
!CC    ANTD
!CC    ANTE
!CC    NOVPT =     Whether vapor pressure value out of temp. range
!CC    NEQN =      Number of Vapor Pressure equation
!CC    TMIN =      Minimum valid temperature for correlation (C)
!CC    TMAX =      Maximum valid temperature for correlation (C)
!CC    ISRC =      Source of the value (DIPPR801 or Yaws)
!CC
!CC Variables Internal to Subroutine VAPORP:
!CC    TEMP =      Temperature for calculation (C for Yaws,
!CC                K for DIPPR801)
!CC
!CC Author:  D. Hokanson (4/3/94)
!CC
!CC****************************************************************

      SUBROUTINE VAPORP(PVAP,TT,ANTA,ANTB,ANTC,ANTD,ANTE,NOVPT,NEQN,TMIN,TMAX,ISRC)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::VAPORP
!MS$ ATTRIBUTES ALIAS:'_VAPORP@48':: VAPORP
!MS$ ATTRIBUTES REFERENCE::PVAP,TT,ANTA,ANTB,ANTC,ANTD,ANTE,NOVPT,NEQN,TMIN,TMAX,ISRC

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION PVAP,TT,ANTA,ANTB,ANTC,ANTD,ANTE,TEMP

         IF (ISRC.EQ.1) THEN
!CC           *** Yaws
            TEMP=TT-273.15D0
         ELSE
!CC           *** DIPPR801
            TEMP=TT
         ENDIF

         IF (NEQN.EQ.101) THEN
!CC           *** DIPPR801
            PVAP=DEXP(ANTA+(ANTB/TEMP)+(ANTC*DLOG(TEMP))+(ANTD*(TEMP**ANTE)))
         ELSE IF (NEQN.LT.0) THEN
!CC           *** Yaws
            PVAP=DEXP(ANTA-ANTB/(TEMP+ANTC))
         ENDIF
         
         TEMP = TT - 273.15D0          
         IF (((TMIN+273.15D0).GT.0).AND.((TEMP.LE.TMIN).OR.(TEMP.GE.TMAX))) THEN
            NOVPT = -1 
         END IF
         
      END

!CC****************************************************************

