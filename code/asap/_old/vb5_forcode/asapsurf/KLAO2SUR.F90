!C***************************************************************
!CC
!CC                           KLAO2SUR
!CC   FIND OXYGEN MASS TRANSFER COEFFICIENT FOR SURFACE AERATION
!CC
!CC Description:  This subroutine calculates the oxygen mass
!CC               transfer coefficient (KLAO2) for surface aeration.
!CC               The correlation to be used gives KLAO2 as a
!CC               function of Power/Volume.
!CC               It comes from the following reference:
!CC
!CC                  Roberts, Paul V. and Paul Dandliker, "Mass
!CC                     Transfer of Volatile Organic Contaminants
!CC                     During Surface Aeration," E.S.&T.,17,8 (1983)
!CC
!CC Output Variable:
!CC    KLAO2 =    Oxygen mass transfer coefficient (1/sec)
!CC
!CC Input Variable:
!CC    POVERV =   Power Input / Unit Volume (W/m^3)
!CC
!C***************************************************************

      SUBROUTINE KLAO2SUR(KLAO2,POVERV)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::KLAO2SUR
!MS$ ATTRIBUTES ALIAS:'_KLAO2SUR':: KLAO2SUR
!MS$ ATTRIBUTES REFERENCE::KLAO2,POVERV

         IMPLICIT DOUBLE PRECISION(A-H,O-Z)
         DOUBLE PRECISION KLAO2,POVERV
   
         KLAO2 = 2.9D-5 * POVERV**0.95D0

      END

!C***************************************************************

