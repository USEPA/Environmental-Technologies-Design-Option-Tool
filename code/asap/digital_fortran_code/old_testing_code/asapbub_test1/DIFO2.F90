!C***************************************************************
!CC
!CC                   DIFO2
!CC           FIND DIFFUSIVITY OF OXYGEN
!CC
!CC Description:  This subroutine calculates the diffusivity of
!CC               oxygen.  The correlation to be used gives
!CC               diffusivity as a function of temperature.
!CC               It comes from the following reference:
!CC
!CC                  Holmen, Kim and Peter Liss, "Models for air-
!CC                     water gas transfer:  an experimental
!CC                     investigation," Tellus 36B (1984).
!CC
!CC Output Variable:
!CC    DIFLO2 =   Diffusivity of oxygen (m^2/sec)
!CC
!CC Input Variable:
!CC    TEMP =     Temperature (K)
!CC
!CC Variables Internal to Subroutine DIFO2
!CC    A,B =      Parameters for fit of data
!CC
!C***************************************************************

      SUBROUTINE DIFO2(DIFLO2,TEMP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::DIFO2
!MS$ ATTRIBUTES ALIAS:'_DIFO2':: DIFO2
!MS$ ATTRIBUTES REFERENCE::DIFLO2,TEMP
   
      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION DIFLO2,TEMP,A,B

         A = 3.15D0
         B = -831.0D0
         DIFLO2 = (10**(A+B/TEMP))*1.0D-9     

      END

!C***************************************************************
 
