!CC*********************************************************************
!CC
!CC                           AQSFIT
!CC
!CC Description:  This subroutine will do a fit of the aqueous
!CC               solubility from the database with the value
!CC               from UNIFAC at the operating T.  It is done by
!CC               using an offset.
!CC
!CC Output Variables:
!CC    AQSOLFIT =      Solubility from UNIFAC Fit (PPMw)
!CC    IAQSOLFITSHT =  Source of this value (Short Version)
!CC    IAQSOLFITLNG =  Source of this value (Long Version)
!CC    IAQSOLFITERR =  Error flag
!CC    AQSOLFITTMP =   Temperature of this value (C)
!CC
!CC Input Variables:
!CC    AQSOLUNDBT =    UNIFAC aqueous solubility at database T (PPMw)
!CC    AQSOLUNDBTTMP = Temperature of above value (C)
!CC    AQSOLUNOPT =    UNIFAC aqueous solubility at operating T (PPMw)
!CC    AQSOL =         Database aqueous solubility (PPMw)
!CC    AQSOLTMP =      Temperature of database aqueous solubility (C)
!CC    TEMPOP =        Operating Temperature (C)
!CC
!CC Variables Internal to Subroutine AQSFIT:
!CC    OFFSET =        Value to offset operating T UNIFAC aqueous
!CC                    solubility by to achieve a fit using a data point
!CC
!CC Author:  D. Hokanson (4/5/94)
!CC
!CC*********************************************************************


      SUBROUTINE AQSFIT(AQSOLFIT,IAQSOLFITSHT,IAQSOLFITLNG,IAQSOLFITERR,AQSOLFITTMP,AQSOLUNDBT,AQSOLUNDBTTMP,AQSOLUNOPT,AQSOL,AQSOLTMP,TEMPOP)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AQSFIT
!MS$ ATTRIBUTES ALIAS:'_AQSFIT':: AQSFIT
!MS$ ATTRIBUTES REFERENCE::AQSOLFIT,IAQSOLFITSHT,IAQSOLFITLNG,IAQSOLFITERR,AQSOLFITTMP,AQSOLUNDBT,AQSOLUNDBTTMP,AQSOLUNOPT,AQSOL,AQSOLTMP,TEMPOP

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)

         IAQSOLFITSHT = 19
      
         OFFSET = AQSOLUNDBT - AQSOL
         AQSOLFIT = AQSOLUNOPT - OFFSET
         AQSOLFITTMP = TEMPOP

      END


