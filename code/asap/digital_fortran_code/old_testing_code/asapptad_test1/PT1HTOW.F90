!C***************************************************************
!CC
!CC                      PT1HTOW
!CC
!CC Description:  This subroutine will find the packed tower
!CC               length necessary to meet design specifications.
!CC
!CC Output Variable:
!CC    HLL =      Tower length (m)
!CC
!CC Input Variables:
!CC    HTU =      Height of a transfer unit (m)
!CC    NTU =      No. of transfer units (-)
!CC
!CC History:  Subroutine written by David R. Hokanson (9/30/93)
!CC
!C***************************************************************

      SUBROUTINE PT1HTOW(HLL,HTU,NTU)
!C  ATTRIBUTES DLLEXPORT, STDCALL::PT1HTOW
!C  ATTRIBUTES ALIAS:'_PT1HTOW@12':: PT1HTOW
!C  ATTRIBUTES REFERENCE::HLL,HTU,NTU

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      DOUBLE PRECISION HLL,HTU,NTU

         HLL = HTU * NTU                                 

      END

!C***************************************************************

