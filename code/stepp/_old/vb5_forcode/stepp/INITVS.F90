!CC*****************************************************************
!CC
!CC                  INITVS
!CC             INITIALIZE VARIABLES
!CC
!CC Description:  This subroutine will initialize the variables needed
!CC               for the STEPP program.
!CC
!CC Output Variables:
!CC    TOL =      Tolerance in Newton-Raphson algorithm
!CC    IMAX =     Maximum number of iterations in Newton-Raphson
!CC    MS =       Array of UNIFAC groups
!CC    NMAX =     Maximum number of UNIFAC groups
!CC    XX =       Initial guess for solution of Newton-Raphson
!CC    NG =
!CC    NDIF =
!CC
!CC Input Variables:
!CC    MX =       Maximum number of UNIFAC groups
!CC
!CC Author:  D. Hokanson (4/4/94)
!CC
!CC*****************************************************************

      SUBROUTINE INITVS(MX)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::INITVS
!MS$ ATTRIBUTES ALIAS:'_INITVS@4':: INITVS
!MS$ ATTRIBUTES REFERENCE::MX

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      PARAMETER (ND = 10)
      COMMON /LIMITS/ TOL, IMAX
      COMMON /GROUP/ MS(10,10,2), NMAX
      COMMON /INIT/ XX(10), NG, NDIF
 
         TOL = 1.0D-10
         IMAX = 250
         MS(1,1,1)=17
         MS(1,1,2)=1 
         NMAX=MX
         XX(1)=1.0D0
         XX(2)=0.0D0
         NG=0 
         NDIF=0 

      END

!CC*****************************************************************

