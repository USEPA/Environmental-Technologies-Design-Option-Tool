!CC********************************************************************
!CC
!CC                      HENFIT
!CC
!CC Description:  This subroutine will calculate Henry's constant by
!CC               fitting a data point from the database to the
!CC               UNIFAC curve.  It makes use of an offset, taking
!CC               advantage of Henry's constant's linear dependence on
!CC               ln(T).  Only one data point is used to calculate the
!CC               offset.  If more than one data point is available,
!CC               the data point used is the one at the temperature
!CC               closest to the operating temperature.
!CC
!CC Output Variables:
!CC    HCFIT =     Henry's constant from UNIFAC fit (dimensionless)
!CC    IHCFITSHT = Source of this value (Short Version)
!CC    IHCFITLNG = Source of this value (Long Version)
!CC    IHCFITERR = Error Flag
!CC    HCFITTMP =  Temperature of Interest (C)
!CC
!CC Input Variables:
!CC    HCDB =        Array of database Henry's constants (-)
!CC    HCDBTMP =     Array of DB Henry's constant temperatures (C)
!CC    HCUNOPT =     UNIFAC Henry's constant value at operating T (-)
!CC    HCUNVAL =     Array of UNIFAC Henry's constants at DB T's (-)
!CC    IHCUNERR =    Array of UNIFAC Henry's constant errors
!CC    TEMPOP =      Operating Temperature (C)
!CC    NUMDBHCS =    Number of Henry's constant data points in DB
!CC
!CC Variables Internal to Subroutine HENFIT:
!CC    CURRDIFF =    Variable used to find closest temp. to operating T
!CC    PERMDIFF =    Same definition as CURRDIFF
!CC    CLOSEHC =     Current Henry's constant closest to operating temp.
!CC    HCDATAPT =    HC in database at T closest to operating T (atm)
!CC    HCDATAPTTMP = Temperature of this point (C)
!CC    HCUNPT =      UNIFAC HC at T closest to database T (atm)
!CC    HCUNPTTMP =   Temperature of this point (C)
!CC    HCUNOPTPT =   UNIFAC HC at operating T (atm)
!CC
!CC Author:  D. Hokanson (4/4/94)
!CC
!CC********************************************************************


      SUBROUTINE HENFIT(HCFIT,IHCFITSHT,IHCFITLNG,IHCFITERR,HCFITTMP,HCDB,HCDBTMP,HCUNOPT,HCUNVAL,IHCUNERR,TEMPOP,NUMDBHCS)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::HENFIT
!MS$ ATTRIBUTES ALIAS:'_HENFIT':: HENFIT
!MS$ ATTRIBUTES REFERENCE::HCFIT,IHCFITSHT,IHCFITLNG,IHCFITERR,HCFITTMP,HCDB,HCDBTMP,HCUNOPT,HCUNVAL,IHCUNERR,TEMPOP,NUMDBHCS

      IMPLICIT DOUBLE PRECISION(A-H,O-Z)
      PARAMETER (NUMHCS = 20)
      DOUBLE PRECISION HCDATAPT,HCDATAPTTMP,HCDB(NUMHCS)
      DOUBLE PRECISION HCDBTMP(NUMHCS),HCUNVAL(NUMHCS)
      DOUBLE PRECISION HCUNPT,HCUNPTTMP,HCUNOPTPT
      DOUBLE PRECISION OFFSET,LNHCFIT
      INTEGER CLOSEHC, IHCUNERR(NUMHCS)
   
         IHCFITSHT = 18
      
!CC*********** find point in database closest to operating temp. if
!CC*********** more than one data point is available

         IF (NUMDBHCS.EQ.1) THEN
            HCDATAPT = HCDB(1)
            HCDATAPTTMP = HCDBTMP(1)
            HCUNPT = HCUNVAL(1)
            HCUNPTTMP = HCDBTMP(1)
         ELSE
            I = 1
 10         IF (IHCUNERR(I).GE.0) THEN
               PERMDIFF = TEMPOP - HCDBTMP(I)
               CLOSEHC = I
               NUM = I+1
               GOTO 15
            ELSE
               I = I + 1
               GOTO 10   
            END IF
 15         DO 20, I=NUM,NUMDBHCS
               IF (IHCUNERR(I).LT.0) GOTO 20
               CURRDIFF = TEMPOP - HCDBTMP(I)
               IF (DABS(CURRDIFF).LT.DABS(PERMDIFF)) THEN
                  CLOSEHC = I
                  PERMDIFF = CURRDIFF
               END IF      
 20         CONTINUE
            HCDATAPT = HCDB(CLOSEHC)
            HCDATAPTTMP = HCDBTMP(CLOSEHC)
            HCUNPT = HCUNVAL(CLOSEHC)
            HCUNPTTMP = HCDBTMP(CLOSEHC)
         END IF

         HCUNOPTPT = HCUNOPT

!CC********* convert values to units of atmospheres

         HCDATAPT = HCDATAPT*(HCDATAPTTMP+273.15D0)*0.082054D0*55.5D0
         HCUNPT = HCUNPT * (HCUNPTTMP+273.15D0) * 0.082054D0 * 55.5D0
         HCUNOPTPT = HCUNOPTPT * (TEMPOP+273.15D0) * 0.082054D0*55.5D0


         OFFSET = DLOG(HCUNPT) - DLOG(HCDATAPT)
         LNHCFIT = DLOG(HCUNOPTPT) - OFFSET
         HCFIT = DEXP(LNHCFIT)

         HCFITTMP = TEMPOP

         HCFIT = HCFIT / 0.082054D0 / (TEMPOP+273.15D0) / 55.5D0

      END


