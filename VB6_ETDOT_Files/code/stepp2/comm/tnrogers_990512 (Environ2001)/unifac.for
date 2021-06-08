$NODEBUG
$NOFLOATCALLS
$STORAGE:4
      PROGRAM  IPES_2
C
C  *********************************************************************
C  *                                                                   *
C  *     << Environmental Properties / Adsorption Isotherms >>         *
C  *                                                                   *
C  *                                                                   *
C  *     DRAFT VERSION: August 18, 1994          Tony N. Rogers        *
C  *                                             S. Bhuvendralingam    *
C  *                                             David W. Hand         *
C  *     TO CONTACT THE AUTHORS:                 John C. Crittenden    *
C  *                                                                   *
C  *     Michigan Technological University                             *
C  *     Dept. of Civil & Environmental Engineering                    *
C  *     Houghton, MI  49931                                           *
C  *     Tel: (906) 487-2210                                           *
C  *                                                                   *
C  *********************************************************************
C
      IMPLICIT REAL*8(A-H,O-Z)
      CHARACTER*80  NAME,NTYPE
      DIMENSION  NHEAD(40)
      PARAMETER  (MA=58, NA=116, NC=2)
      COMMON /A/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
      COMMON /F/ TOL,IMAX
      COMMON /G/ MS(10,10,2),NMAX
      COMMON /I/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
      COMMON /O/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
      COMMON /V/ NEQN,ANTA,ANTB,ANTC,ANTD,ANTE,TMIN,TMAX,PS,CS
C
C        -- SET UP I/O FILE NAMES --
C
      OPEN(5,FILE='IPES.DAT',STATUS='UNKNOWN',
     &       ACCESS='SEQUENTIAL',FORM='FORMATTED')
      OPEN(4,FILE='TABLE.OUT',STATUS='UNKNOWN',
     &       ACCESS='SEQUENTIAL',FORM='FORMATTED')
      ENDFILE 4
      REWIND 4
      OPEN(7,FILE='LABEL.OUT',STATUS='UNKNOWN',
     &       ACCESS='SEQUENTIAL',FORM='FORMATTED')
      ENDFILE 7
      REWIND 7
      OPEN(8,FILE='IOWN.DAT',STATUS='UNKNOWN',
     &       ACCESS='SEQUENTIAL',FORM='FORMATTED')
      OPEN(11,FILE='PROPS.OUT',STATUS='UNKNOWN',
     &        ACCESS='SEQUENTIAL',FORM='FORMATTED')
      ENDFILE 11
      REWIND 11
C     OPEN(12,FILE='SAMPLE.OUT',STATUS='UNKNOWN',
C    &        ACCESS='SEQUENTIAL',FORM='FORMATTED')
C     ENDFILE 12
C     REWIND 12
C
C        -- SAMPLE INPUT DATA (BENZENE ON CALGON F-300 CARBON) --
C
      IMOD = 1
      MDL = 1
      IMAX = 250
      NTYPE = 'Calgon F-300 carbon'
      W0 = 1.72D0
      BB = 7.42D-02
      GM = 1.D0
      RELHUM = 0.D0
C
      NAME = 'BENZENE'
      ID = 1
      TC = 25.D0
      FWT = 0.D0
      SOLUB = 0.D0
      VBM = 0.D0
      RNDX = 0.D0
      ANTA = 0.D0
      ANTB = 0.D0
      ANTC = 0.D0
      CBULK = 50.D0
      OMAG = 4.D0
C
C     ... "NL" is the initial number of integration points in "SPEQ",
C         as well as the number of regression points in the
C         Freundlich isotherm calculations; "TOL" is a user-
C         supplied convergence tolerance (for the solubility and
C         spreading pressure algorithms) ...
C
      NL = 25
      TOL = 1.D-06
      XERR = 0.1D0
      SPRD = 0.D0
C
      NSOL = 0
      NVOL = 0
      IRNG = 1
      NMAX = 1
      MS(NC,1,1) = 10
      MS(NC,1,2) = 6
C
C        -- READ INPUT FILE --
C
         READ(5,910) NHEAD
         READ(5,*) IMOD,MDL,IMAX,TOL
         READ(5,910) NHEAD
         READ(5,912) NTYPE
         READ(5,910) NHEAD
         READ(5,*) W0,BB,GM,RELHUM
         READ(5,910) NHEAD
      IF((IMOD.LT.0).OR.(IMOD.GT.6)) IMOD=0
      IF((MDL.LT.1).OR.(MDL.GT.3)) MDL=1
      IF(RELHUM.LT.60.D0) THEN RELHUM=0.D0
C
C        -- PHYSICAL PROPERTY "HEADER" (IMOD=0) --
C
      WRITE(11,'(//)')
      IF(MDL.EQ.1) THEN
         WRITE(11,901)
         WRITE(*,901)
      ELSEIF(MDL.EQ.2) THEN
         WRITE(11,902)
         WRITE(*,902)
      ELSEIF(MDL.EQ.3) THEN
         WRITE(11,903)
         WRITE(*,903)
      ENDIF
      WRITE(11,911)
      WRITE(11,914)
      IF(IMOD.EQ.0) GOTO 3
C
C        -- ISOTHERM "HEADER" (IMOD>0) --
C
      WRITE(4,900)
      WRITE(7,900)
      WRITE(*,900)
      IF(IMOD.EQ.1) THEN
         WRITE(4,'(5X,A,/,5X,A,//)')
     &   'LIQUID PHASE : 3-Parameter Polanyi Correlation',
     &   '               [ Uniform adsorbate ]'
         WRITE(7,'(5X,A,/,5X,A,//)')
     &   'LIQUID PHASE : 3-Parameter Polanyi Correlation',
     &   '               [ Uniform adsorbate ]'
         WRITE(*,'(5X,A,/,5X,A,//)')
     &   'LIQUID PHASE : 3-Parameter Polanyi Correlation',
     &   '               [ Uniform adsorbate ]'
         WRITE(4,909) NTYPE,W0,BB,GM
         WRITE(7,909) NTYPE,W0,BB,GM
         WRITE(*,909) NTYPE,W0,BB,GM
      ELSEIF(IMOD.EQ.2) THEN
         WRITE(4,'(5X,A,//,5X,A,//)')
     &   'VAPOR PHASE : Dubinin-Radushkevich (D-R) Correlation',
     &   '              [ Relative Humidity = 0 % ]'
         WRITE(7,'(5X,A,//,5X,A,//)')
     &   'VAPOR PHASE : Dubinin-Radushkevich (D-R) Correlation',
     &   '              [ Relative Humidity = 0 % ]'
         WRITE(*,'(5X,A,//,5X,A,//)')
     &   'VAPOR PHASE : Dubinin-Radushkevich (D-R) Correlation',
     &   '              [ Relative Humidity = 0 % ]'
      ELSEIF(IMOD.EQ.3) THEN
         WRITE(4,'(5X,A,/,5X,A,//,5X,A,0PF5.1,A,//)')
     &   'VAPOR PHASE : Calgon BPL "Characteristic Curve" Correlation,',
     &   '              Grant-Joyce-Urbanic correction for R.H. > 60 %',
     &   '              [ Relative Humidity = ',RELHUM,' % ]'
         WRITE(7,'(5X,A,/,5X,A,//,5X,A,0PF5.1,A,//)')
     &   'VAPOR PHASE : Calgon BPL "Characteristic Curve" Correlation,',
     &   '              Grant-Joyce-Urbanic correction for R.H. > 60 %',
     &   '              [ Relative Humidity = ',RELHUM,' % ]'
         WRITE(*,'(5X,A,/,5X,A,//,5X,A,0PF5.1,A,//)')
     &   'VAPOR PHASE : Calgon BPL "Characteristic Curve" Correlation,',
     &   '              Grant-Joyce-Urbanic correction for R.H. > 60 %',
     &   '              [ Relative Humidity = ',RELHUM,' % ]'
      ELSEIF(IMOD.EQ.4) THEN
         WRITE(4,'(5X,A,//,5X,A,//)')
     &  'VAPOR PHASE : D-R Equal Spreading Pressure (SPEQ) Calculation',
     &  '              [ Relative Humidity = 0 % ]'
         WRITE(7,'(5X,A,//,5X,A,//)')
     &  'VAPOR PHASE : D-R Equal Spreading Pressure (SPEQ) Calculation',
     &  '              [ Relative Humidity = 0 % ]'
         WRITE(*,'(5X,A,//,5X,A,//)')
     &  'VAPOR PHASE : D-R Equal Spreading Pressure (SPEQ) Calculation',
     &  '              [ Relative Humidity = 0 % ]'
      ELSEIF(IMOD.EQ.5) THEN
         WRITE(4,'(5X,A,//,5X,A,//)')
     &   'LIQUID PHASE : Estimated from Gas-Phase D-R Isotherm',
     &   '               [ Manes-Hofer, uniform adsorbate ]'
         WRITE(7,'(5X,A,//,5X,A,//)')
     &   'LIQUID PHASE : Estimated from Gas-Phase D-R Isotherm',
     &   '               [ Manes-Hofer, uniform adsorbate ]'
         WRITE(*,'(5X,A,//,5X,A,//)')
     &   'LIQUID PHASE : Estimated from Gas-Phase D-R Isotherm',
     &   '               [ Manes-Hofer, uniform adsorbate ]'
      ELSEIF(IMOD.EQ.6) THEN
         WRITE(4,'(5X,A,/,5X,A,//)')
     &   'LIQUID PHASE : Estimated from Gas-Phase D-R Isotherm',
     &   '               [ Hansen-Fackler, non-uniform adsorbate ]'
         WRITE(7,'(5X,A,/,5X,A,//)')
     &   'LIQUID PHASE : Estimated from Gas-Phase D-R Isotherm',
     &   '               [ Hansen-Fackler, non-uniform adsorbate ]'
         WRITE(*,'(5X,A,/,5X,A,//)')
     &   'LIQUID PHASE : Estimated from Gas-Phase D-R Isotherm',
     &   '               [ Hansen-Fackler, non-uniform adsorbate ]'
         WRITE(4,909) NTYPE,W0,BB,GM
         WRITE(7,909) NTYPE,W0,BB,GM
         WRITE(*,909) NTYPE,W0,BB,GM
      ENDIF
      IF((IMOD.LT.2).OR.(IMOD.GT.4)) THEN
         IF(MDL.EQ.1) THEN
            WRITE(4,901)
            WRITE(7,901)
            WRITE(*,901)
         ELSEIF(MDL.EQ.2) THEN
            WRITE(4,902)
            WRITE(7,902)
            WRITE(*,902)
         ELSEIF(MDL.EQ.3) THEN
            WRITE(4,903)
            WRITE(7,903)
            WRITE(*,903)
         ENDIF
      ENDIF
      WRITE(4,911)
      WRITE(4,913)
      WRITE(4,911)
      IF(IMOD.NE.4) THEN
         WRITE(4,906)
      ELSE
         WRITE(4,919)
      ENDIF
C
C        -- LOAD "UNIFAC" BINARY INTERACTION PARAMETERS --
C
   3  CONTINUE
      CALL BINPAR(MDL,MGSG,AI,RI,QI,FMW,FVB)
C
C        -- ITERATE OVER ENTIRE CHEMICAL LIST --
C
   5  CONTINUE
         NAME=' '
         DO 10 I=1,NC
         DO 10 J=1,10
         DO 10 K=1,2
            MS(I,J,K)=0
  10     CONTINUE
         READ(5,*,ERR=20) NAME,ID,TC,FWT,HLQ,SOLUB,XKQ,VBM,RNDX,
     &                    NEQN,ANTA,ANTB,ANTC,ANTD,ANTE,TMIN,TMAX,
     &                    CBULK,OMAG,NL,NSOL,NVOL,IRNG,
     &                    NMAX,(MS(NC,K,1),MS(NC,K,2),K=1,NMAX)
      IF(ID.LE.0) GOTO 15
      TT=TC+273.15D0
      MS(1,1,1)=17
      MS(1,1,2)=1
C
C        >> PHYSICAL PROPERTIES:
C
      HLC=0.D0
      GAMMA=0.D0
      PVAP=0.D0
      XXFWT=0.D0
      XXSOL=0.D0
      XXTIE=0.D0
      XXVBM=0.D0
      NPRNT=0
      CALL HENRY(TT,HLC,GAMMA,PVAP,NPRNT,MERR)
      NPRNT=0
      KERR=0
      CALL AQSOL(0,0,IRNG,NC,TT,XXFWT,XXSOL,XXTIE,XXVBM,NPRNT,KERR)
C
C        -- OCTANOL AND WATER MOLAR DENSITIES, [GMOL/L] --
C
      OCTDEN=6.36D0
      WATDEN=55.5D0
      CALL PARTC(TT,OCTDEN,WATDEN,XKOW,XLGK,IERR)
C
C        >> ESTIMATE LIQUID DENSITY OF CHEMICAL:
C           (Reference chemical is water at T of interest)
C
      ORGDEN=0.D0
      IF((XXFWT.GT.TOL).AND.(XXVBM.GT.TOL)) THEN
         PW =  0.95D0
         A1 = -1.4176800403D+00
         A2 =  8.9766515240D+00
         A3 = -1.2275501969D+01
         A4 =  7.4584410413D+00
         A5 = -1.7384916050D+00
         XAVG = 324.65D+00
         FAVG = 0.98396D+00
         XN = TT/XAVG
         FN = A1 + A2*(XN) + A3*(XN)**2 + A4*(XN)**3 + A5*(XN)**4
         FX = FN*FAVG
         ORGDEN = PW*FX*(XXFWT/XXVBM)/(18.015D0/21.D0)
      ENDIF
C
C        >> ESTIMATES OF HENRY'S CONSTANT AND SOLUBILITY:
C
      HHEST=0.D0
      SSEST=0.D0
      IF((PVAP.GT.TOL).AND.(XXSOL.GT.TOL).AND.(XXFWT.GT.TOL)) THEN
         HHEST=(PVAP/760.D0)/(XXSOL/XXFWT/1.D03)/(0.08206D0*TT)
      ENDIF
      IF((GAMMA.GT.TOL).AND.(ORGDEN.GT.TOL).AND.(XXFWT.GT.TOL)) THEN
         XS=1.D0/GAMMA
         AVG=(XS*XXFWT)+((1.D0-XS)*18.015D0)
         XF=(XS*XXFWT)/AVG
         DH2O=FX
         DENS=1.D0/(XF/ORGDEN+(1.D0-XF)/DH2O)
         SSEST=XF*DENS*1.D06
      ENDIF
C
      WRITE(11,915) TC,ID,XXFWT,GAMMA,HLC,XXSOL,XXTIE,XKOW,XLGK,
     &              XXVBM,ORGDEN,PVAP,HHEST,SSEST,NAME
      IF(NPRNT.NE.0) THEN
         IF(MERR.EQ.-1) WRITE(11,916)
         IF(KERR.EQ.-1) WRITE(11,917)
         IF(IERR.EQ.-1) WRITE(11,918)
         IF((MERR.EQ.-1).OR.(KERR.EQ.-1).OR.(IERR.EQ.-1)) THEN
            WRITE(11,904) ID,NAME
         ENDIF
      ENDIF
      IF(IMOD.EQ.0) GOTO 15
C
C        >> ADSORPTION ISOTHERM MODELS:
C
      WRITE(7,911)
      WRITE(*,911)
      IF(IMOD.EQ.1) THEN
         CALL ADLIQ(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,VBM,KERR)
      ELSEIF((IMOD.EQ.2).OR.(IMOD.EQ.3)) THEN
         CALL ADVAP(RELHUM,IMOD,NVOL,IRNG,NC,TT,FWT,VBM,RNDX,KERR)
         SOLUB=CS
      ELSEIF(IMOD.EQ.4) THEN
         I1=2
         I2=NVOL
         I3=IRNG
         I4=NC
         I5=NL
         R1=0.D0
         R2=TT
         R3=FWT
C     PRINT *,'FWT = ',R3
         R4=VBM
         R5=RNDX
         R6=OMAG
         CALL ADVAP(R1,I1,I2,I3,I4,R2,R3,R4,R5,KERR)
C        CALL ADVAP(0.D0,2,NVOL,IRNG,NC,TT,FWT,VBM,RNDX,KERR)
C     PRINT *,'FWT = ',R3
         IF(KERR.EQ.-1) GOTO 12
         IF(XNF.GT.1.D0) THEN
            KERR=-1
            WRITE(4,922)
            WRITE(7,922)
            WRITE(*,922)
            GOTO 12
         ENDIF
         NL=I5
         OMAG=R6
         CALL SPEQ(NVOL,IRNG,NC,TT,FWT,VBM,RNDX,SPRD,BETA,XERR,KERR)
C     PRINT *,'FWT = ',R3
         SOLUB=CS
      ELSEIF(IMOD.EQ.5) THEN
         CALL HOFMAN(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,VBM,RNDX,KERR)
      ELSEIF(IMOD.EQ.6) THEN
         CALL HANFAC(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,VBM,RNDX,KERR)
      ENDIF
  12  CONTINUE
      IF(KERR.EQ.-1) THEN
         WRITE(4,904) ID,NAME
         WRITE(7,904) ID,NAME
         WRITE(*,904) ID,NAME
         GOTO 15
      ENDIF
      IF(IMOD.NE.4) THEN
         WRITE(4,907)
     &   TC,ID,FWT,SOLUB,ORGDEN,XK1,XK2,XNF,CSAV,QSAV,
     &   CBEG,CEND,NL,RSQD,RMSE,NAME
         WRITE(7,905)
     &   TC,ID,NAME,FWT,ORGDEN,SOLUB,CSAV,QSAV,XK1,XK2,XNF,
     &   CBEG,CEND,NL,RSQD,RMSE
         WRITE(*,905)
     &   TC,ID,NAME,FWT,ORGDEN,SOLUB,CSAV,QSAV,XK1,XK2,XNF,
     &   CBEG,CEND,NL,RSQD,RMSE
      ELSE
         WRITE(4,920)
     &   TC,ID,FWT,SOLUB,ORGDEN,XK1,XK2,XNF,SPRD,BETA,CSAV,QSAV,
     &   CBEG,CEND,NL,NAME
         WRITE(7,921)
     &   TC,ID,NAME,FWT,ORGDEN,BETA,SOLUB,CSAV,QSAV,SPRD,XK1,XK2,XNF,
     &   CBEG,CEND,NL,(XERR*100.D0)
         WRITE(*,921)
     &   TC,ID,NAME,FWT,ORGDEN,BETA,SOLUB,CSAV,QSAV,SPRD,XK1,XK2,XNF,
     &   CBEG,CEND,NL,(XERR*100.D0)
      ENDIF
  15  CONTINUE
      GOTO 5
  20  CONTINUE
C
      IF(IMOD.NE.0) THEN
         WRITE(7,911)
         WRITE(*,911)
         WRITE(*,908)
      ENDIF
      CLOSE(4)
      CLOSE(5)
      CLOSE(7)
      CLOSE(8)
      CLOSE(11)
      CLOSE(12)
      STOP
C
C        -- FORMAT STATEMENTS --
C
 900  FORMAT(//,2X,
     & '** POLANYI ESTIMATION OF FREUNDLICH ISOTHERM PARAMETERS **',//)
 901  FORMAT(5X,'UNIFAC Database : Original VLE',/)
 902  FORMAT(5X,'UNIFAC Database : Original LLE',/)
 903  FORMAT(5X,'UNIFAC Database : Environmental VLE',/)
 904  FORMAT(2X,'=> Calculation aborted for component ID #',I3,
     &          ', ',A,/)
 905  FORMAT(16X,'           Temperature : ',0PF7.1,' [C]',/,
     & 9X,' Adsorbate chemical (ID #',I3,') :  ',A,/,
     & 9X,' Molecular weight of chemical : ',0PF7.2,/,
     & 9X,'   Liquid density of chemical : ',0PF7.4,' [g/cc]',/,
     & 9X,'     Saturation concentration : ',1PE11.4,' [ppmw]',//,
     & 9X,'  Avg. chemical concentration : ',1PE11.4,' [ppbw]',/,
     & 9X,'  Polanyi adsorption capacity : ',1PE11.4,
     &    ' [ug/gm carbon]',//,
     & 9X,'              Freundlich "K1" : ',1PE11.4,
     &    ' [ug/gm (L/ug)**1/n]',/,
     & 9X,'              Freundlich "K2" : ',1PE11.4,
     &    ' [umol/gm (L/umol)**1/n]',/,
     & 9X,'             Freundlich "1/n" : ',0PF7.4,//,
     & 9X,'      Lower correlation limit : ',1PE11.4,' [ppbw]',/,
     & 9X,'      Upper correlation limit : ',1PE11.4,' [ppbw]',/,
     & 9X,'  Number of regression points : ',3X,I4,/,
     & 9X,'         Regression r-squared : ',0PF7.4,/,
     & 9X,'Root-Mean-Square Error (RMSE) : ',0PF7.2,' [%]',/)
 906  FORMAT(5X,'T-C',4X,'ID',5X,'MWT',8X,'SAT',8X,'RHO',
     &       9X,'K1',9X,'K2',8X,'1/N',7X,'C-AVG',6X,'Q-AVG',
     &       6X,'C-BEG',6X,'C-END',4X,'N-REGR',3X,'R^2',6X,'RMSE',
     &       5X,'CHEMICAL NAME:',/)
 907  FORMAT(2X,0PF6.1,3X,I3,2X,0PF7.2,2X,4(1PE11.3),1X,0PF8.4,
     &       2X,4(1PE11.3),I6,2X,0PF7.4,2X,0PF7.2,5X,A)
 908  FORMAT(//,2X,
     & '** "ASCII" FILES: "TABLE.OUT", "LABEL.OUT", AND "PROPS.OUT"',//)
 909  FORMAT(5X,'Polanyi Isotherm,  W=Wo*exp{-B*(e/v)^G},  for ',A,//,
     &      10X,'Total pore volume, Wo  = ',1PE10.3,
     &          '  [cc/gm carbon]',/,
     &      10X,'Polanyi parameter "B"  = ',1PE10.3,/,
     &      10X,'Polanyi exponent, "G"  = ',1PE10.3,//)
 910  FORMAT(40A2)
 911  FORMAT(/,2X,75('-'),//)
 912  FORMAT(A80)
 913  FORMAT(9X,'* OUTPUT KEY TO VARIABLES AND UNITS *',//,
     & '     T-C = Temperature in [C]',/,
     & '      ID = ID Number of adsorbed chemical',/,
     & '     FWT = Molecular weight of chemical',/,
     & '     SAT = Saturation concentration in [ppmw]',/,
     & '     RHO = Liquid density of chemical in [g/cc]',/,
     & '      K1 = Freundlich "K" in [ug/gm (L/ug)**1/n]',/,
     & '      K2 = Freundlich "K" in [umol/gm (L/umol)**1/n]',/,
     & '     1/N = Freundlich "1/n"',/,
     & '   SP-PR = Integrated spreading pressure',/,
     & '   POLAR = Normalization factor "BETA"',/,
     & '   C-AVG = Avg. chemical concentration in [ppbw]',/,
     & '   Q-AVG = Adsorption capacity in [ug/gm carbon]',/,
     & '   C-BEG = Lower correlation limit in [ppbw]',/,
     & '   C-END = Upper correlation limit in [ppbw]',/,
     & '  N-REGR = Number of regression points',/,
     & '    RSQD = Regression r-squared',/,
     & '    RMSE = Root-Mean-Square Error as [%]',/)
 914  FORMAT(3X,'* PHYSICAL PROPERTY TABULATION *',///,
     &       5X,'T-C',4X,'ID',5X,'MWT',7X,'GAMMA',7X,'HLC',7X,'SOLUB',
     &       7X,'TIE',8X,'KOW',7X,'LOG_K',6X,'VOL_M',5X,'DENSITY',
     &       5X,'P_VAP',6X,'H_EST',6X,'S_EST',7X,'CHEMICAL NAME:',/,
     &      40X,'{dim}',6X,'{ppmw}',5X,'{ppmw}',26X,'{cc/mol}',
     &       4X,'{g/cc}',4X,'{mm Hg}',5X,'{dim}',6X,'{ppmw}',/)
 915  FORMAT(2X,0PF6.1,3X,I3,2X,0PF7.2,2X,11(1PE11.3),5X,A)
 916  FORMAT(2X,'** ERROR: Problem with value of Henry"s Constant **')
 917  FORMAT(2X,'** ERROR: Problem with value of solubility **')
 918  FORMAT(2X,'** ERROR: Problem with value of K,ow coefficient **')
 919  FORMAT(5X,'T-C',4X,'ID',5X,'MWT',8X,'SAT',8X,'RHO',
     &       9X,'K1',9X,'K2',8X,'1/N',7X,'SP-PR',6X,'POLAR',
     &       6X,'C-AVG',6X,'Q-AVG',6X,'C-BEG',6X,'C-END',4X,'N-STEP',
     &       5X,'CHEMICAL NAME:',/)
 920  FORMAT(2X,0PF6.1,3X,I3,2X,0PF7.2,2X,4(1PE11.3),1X,0PF8.4,
     &       2X,6(1PE11.3),2X,I6,5X,A)
 921  FORMAT(16X,'           Temperature : ',0PF7.1,' [C]',/,
     & 9X,' Adsorbate chemical (ID #',I3,') :  ',A,/,
     & 9X,' Molecular weight of chemical : ',0PF7.2,/,
     & 9X,'   Liquid density of chemical : ',0PF7.4,' [g/cc]',/,
     & 9X,'  Normalization factor "BETA" : ',0PF7.4,/,
     & 9X,'     Saturation concentration : ',1PE11.4,' [ppmw]',//,
     & 9X,'  Bulk chemical concentration : ',1PE11.4,' [ppbw]',/,
     & 9X,'      D-R adsorption capacity : ',1PE11.4,
     &    ' [ug/gm carbon]',/,
     & 9X,'Integrated spreading pressure : ',1PE11.4,//,
     & 9X,'              Freundlich "K1" : ',1PE11.4,
     &    ' [ug/gm (L/ug)**1/n]',/,
     & 9X,'              Freundlich "K2" : ',1PE11.4,
     &    ' [umol/gm (L/umol)**1/n]',/,
     & 9X,'             Freundlich "1/n" : ',0PF7.4,//,
     & 9X,'      Lower correlation limit : ',1PE11.4,' [ppbw]',/,
     & 9X,'      Upper correlation limit : ',1PE11.4,' [ppbw]',/,
     & 9X,' Number of integration points : ',1X,I6,/,
     & 9X,'    Error tolerance (vs. D-R) : ',0PF7.1,' [%]',/)
 922  FORMAT(2X,'** ERROR: Freundlich exponent is greater than "1" **',
     &     /,2X,'          SPEQ bypassed -- unfavorable adsorption')
      END
C
C  *********************************************************************
C  *                                                                   *
C  *          << SUBROUTINE TO HANDLE FUNCTIONAL GROUPS >>             *
C  *                                                                   *
C  *********************************************************************
      SUBROUTINE FGRP(MODEL,NSOL,NVOL,IRNG,NC,NG,XMW,FWT,VBM,NPRNT,JERR)
C  *********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      PARAMETER  (MA=58, NA=116)
      DIMENSION  NGM(10),NY(10,20),JH(NA),IH(20),XMW(10),XTW(10)
      COMMON /U/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),
     &           R(10),P(10,10)
      COMMON /A/ AI(MA,MA),RI(NA),QI(NA),FMW(NA),FVB(NA),MGSG(NA)
      COMMON /F/ TOL,IMAX
      COMMON /G/ MS(10,10,2),NMAX
C
      JERR=0
      NK=NC
C
C        -- COMPONENT MOLECULAR WEIGHTS --
C
      DO 103 KO=1,3
         XMW(KO)=0.D0
 103  CONTINUE
         XMW(1)=18.015D0
         XMW(NK)=FWT
      DO 107 LP=1,NK
         XTW(LP)=0.D0
         DO 105 JP=1,NMAX
            IDG=MS(LP,JP,1)
            IF(IDG.EQ.0) GOTO 105
            XTS=FMW(IDG)
            IF(XTS.LE.0.D0) GOTO 107
            XTW(LP)=XTW(LP)+XTS*DBLE(MS(LP,JP,2))
 105     CONTINUE
         IF(XTW(LP).GT.TOL) XMW(LP)=XTW(LP)
 107  CONTINUE
C     print *,'i am here',fwt
      FWT=XMW(NK)
C     print *,'i am here',fwt
      IF(FWT.LE.TOL) GOTO 363
C
C        -- MOLAR VOLUME IN [CC/GMOLE] --
C
      IF(NVOL.EQ.0) THEN
         VTM=0.D0
         DO 108 KJ=1,NMAX
            IDG=MS(NK,KJ,1)
            IF(IDG.EQ.0) GOTO 108
            XTS=FVB(IDG)
            IF(XTS.LE.0.D0) GOTO 109
            VTM=VTM+XTS*DBLE(MS(NK,KJ,2))
 108     CONTINUE
         VBM=VTM
         IF((NMAX.EQ.1).AND.(MS(NK,NMAX,2).EQ.1)) GOTO 109
         VBM=VBM-(DBLE(IRNG)*7.D0)
 109     CONTINUE
      ENDIF
      IF(VBM.LE.TOL) GOTO 365
      IF(NSOL.NE.0) GOTO 265
C
      DO 10 I=1,10
      DO 10 J=1,NK
         QT(I,J)=0.D0
         RT(I,J)=0.D0
  10  CONTINUE
C
      IF(MODEL.NE.1) GOTO 30
         NG=NK
      DO 20 I=1,NK
         READ(8,*) RT(I,I),QT(I,I),(P(I,J),J=1,NK)
  20  CONTINUE
         GOTO 250
  30  CONTINUE
C
      READ(8,*) IOWNRQ,IOWNP
      IF(IOWNRQ.EQ.0) GOTO 50
C
      DO 40 I=1,IOWNRQ
         READ(8,*) K,RI(K),QI(K)
  40  CONTINUE
C
  50  CONTINUE
      IF(IOWNP.EQ.0) GOTO 70
      DO 60 I=1,IOWNP
         READ(8,*) J,K,AI(J,K)
  60  CONTINUE
  70  CONTINUE
C
      DO 90 I=1,NA
         JH(I)=0
  90  CONTINUE
C
      IC=1
      DO 160 I=1,NK
         DO 150 J=1,NMAX
            IF(MS(I,J,1).EQ.0) GOTO 160
            IH(IC)=MS(I,J,1)
            IF(IC.EQ.1) GOTO 140
            IF(IH(IC).EQ.IH(IC-1)) GOTO 150
            IF(IH(IC).GT.IH(IC-1)) GOTO 140
            IF(IC.GT.2) GOTO 110
            IHH=IH(1)
            IH(1)=IH(2)
            IH(2)=IHH
            GOTO 140
 110        I1=IC-1
            DO 130 I2=1,I1
               IF(IH(IC).GT.IH(I2)) GOTO 130
               IF(IH(IC).EQ.IH(I2)) GOTO 150
               I4=IC-I2
               DO 120 I3=1,I4
                  IH(IC+1-I3)=IH(IC-I3)
 120           CONTINUE
               IH(I2)=MS(I,J,1)
 130        CONTINUE
 140        IC=IC+1
            IF(IC.GT.20) GOTO 350
 150     CONTINUE
 160  CONTINUE
C
      IC=IC-1
      DO 170 I=1,IC
         JH(IH(I))=I
 170  CONTINUE
C
      DO 180 I=1,10
      DO 180 J=1,20
         NY(I,J)=0
 180  CONTINUE
C
      DO 200 I=1,NK
         DO 190 J=1,10
            IF(MS(I,J,1).EQ.0) GOTO 200
            N1=MS(I,J,1)
            N2=MS(I,J,2)
            IF(N1.EQ.0) GOTO 200
            N3=JH(N1)
            NY(I,N3)=N2
 190     CONTINUE
 200  CONTINUE
C
      I=0
      NGMGL=0
      DO 210 K=1,IC
         NSG=IH(K)
         NGMNY=MGSG(NSG)
         IF(NGMNY.NE.NGMGL) I=I+1
         NGM(I)=NGMNY
         NGMGL=NGMNY
      DO 210 J=1,NK
         RT(I,J)=RT(I,J)+NY(J,K)*RI(NSG)
         QT(I,J)=QT(I,J)+NY(J,K)*QI(NSG)
 210  CONTINUE
         NG=I
C
      DO 220 I=1,NG
      DO 220 J=1,NG
         NI=NGM(I)
         NJ=NGM(J)
         AVAL=AI(NI,NJ)
         IF(DABS(AVAL).GT.(9.D+04)) GOTO 360
         P(I,J)=AVAL
 220  CONTINUE
 250  CONTINUE
C
      DO 260 I=1,NK
         Q(I)=0.D0
         R(I)=0.D0
      DO 260 K=1,NG
         Q(I)=Q(I)+QT(K,I)
         R(I)=R(I)+RT(K,I)
 260  CONTINUE
C
C        -- ERROR MESSAGES --
C
 265  CONTINUE
         JERR=1
         GOTO 380
 350  CONTINUE
         JERR=-1
         IF(NPRNT.NE.0) THEN
            WRITE(4,901)
            WRITE(7,901)
            WRITE(*,901)
         ENDIF
         GOTO 370
 360  CONTINUE
         JERR=-1
         IF(NPRNT.NE.0) THEN
            WRITE(4,902)
            WRITE(7,902)
            WRITE(*,902)
         ENDIF
         GOTO 370
 363  CONTINUE
         JERR=-1
         IF(NPRNT.NE.0) THEN
            WRITE(4,903)
            WRITE(7,903)
            WRITE(*,903)
         ENDIF
         GOTO 370
 365  CONTINUE
         JERR=-1
         IF(NPRNT.NE.0) THEN
            WRITE(4,904)
            WRITE(7,904)
            WRITE(*,904)
         ENDIF
 370  CONTINUE
C
 380  CONTINUE
      REWIND 8
      RETURN
C
C        -- FORMAT STATEMENTS --
C
 901  FORMAT(2X,'** ERROR: No. of sub-Groups cannot exceed 20 **')
 902  FORMAT(2X,'** ERROR: 1 or more UNIFAC parameters are missing **')
 903  FORMAT(2X,'** ERROR: Problem with value of molecular weight **')
 904  FORMAT(2X,'** ERROR: Problem with value of molar volume **')
      END
C
C  *********************************************************************
C  *                                                                   *
C  *       << SUBROUTINE TO CALCULATE ACTIVITY COEFFICIENTS >>         *
C  *                                                                   *
C  *********************************************************************
      SUBROUTINE UNIMOD(NDIF,NACT,NC,NG,T,X,ACT,DACT,TACT)
C  *********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      COMMON /U/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),
     &           R(10),P(10,10)
      DIMENSION  X(10),GAM(10),ACT(10),DACT(10,10),THETA(10),
     &           PHI(10),RI(10),QI(10),QIL(10),RIL(10),
     &           QID(10),ETAL(10),TACT(10),U(10,10),V(10,10),
     &           DETA(10),DS(10,10),ETA(10),TETAR(10),H3(10,10)
C
         ZCOORD=10.D0
         THETS=0.D0
         PHS=0.D0
      DO 10 I=1,NC
         THETA(I)=X(I)*Q(I)
         PHI(I)=R(I)*X(I)
         THETS=THETS+THETA(I)
         PHS=PHS+PHI(I)
  10  CONTINUE
C
      DO 20 I=1,NC
         THETA(I)=THETA(I)/THETS
         PHI(I)=PHI(I)/PHS
         RI(I)=R(I)/PHS
         RIL(I)=DLOG(RI(I))
         QI(I)=Q(I)/THETS
         QID(I)=1.D0-RI(I)/QI(I)
         QIL(I)=DLOG(QI(I))
  20  CONTINUE
C
      DO 30 I=1,NC
         XX=F(I)+Q(I)*(1.D0-QIL(I))-RI(I)+RIL(I)
         XX=XX-(ZCOORD/2.D0)*Q(I)*(QID(I)+RIL(I)-QIL(I))
         GAM(I)=XX
  30  CONTINUE
C
      DO 50 I=1,NG
         TETAR(I)=0.D0
         ETA(I)=0.D0
         DO 40 J=1,NC
            ETA(I)=ETA(I)+S(I,J)*X(J)
            TETAR(I)=TETAR(I)+QT(I,J)*X(J)
  40     CONTINUE
         ETAL(I)=DLOG(ETA(I))
  50  CONTINUE
C
      DO 70 I=1,NC
         DO 60 J=1,NG
            U(J,I)=S(J,I)/ETA(J)
            V(J,I)=U(J,I)*TETAR(J)
            GAM(I)=GAM(I)-V(J,I)-QT(J,I)*ETAL(J)
  60     CONTINUE
         ACT(I)=DEXP(GAM(I))
         IF(NACT.EQ.1) ACT(I)=ACT(I)*X(I)
  70  CONTINUE
C
      IF(NDIF.EQ.0) GOTO 160
      IF(NDIF.EQ.2) GOTO 110
C
      DO 90 I=1,NC
      DO 90 J=1,NC
         XX=Q(I)*QI(J)*(1.D0-(ZCOORD/2.D0)*QID(I)*QID(J))+(1.D0-RI(I))*
     &      (1.D0-RI(J))
         DO 80 K=1,NG
            XX=XX+U(K,I)*(V(K,J)-QT(K,J))-U(K,J)*QT(K,I)
  80     CONTINUE
         DACT(I,J)=XX
         DACT(J,I)=XX
         IF(NACT.EQ.1) GOTO 90
         DACT(I,J)=DACT(I,J)*ACT(I)
         IF(J.EQ.I) GOTO 90
         DACT(J,I)=DACT(J,I)*ACT(J)
  90  CONTINUE
C
      IF(NACT.EQ.0) GOTO 110
C
      DO 100 I=1,NC
      DO 100 J=1,NC
         DACT(I,J)=ACT(I)*(DACT(I,J)-1.D0)
         IF(J.EQ.I) DACT(I,J)=DACT(I,J)+DEXP(GAM(I))
 100  CONTINUE
 110  CONTINUE
C
      IF(NDIF.EQ.1) GOTO 160
C
      DO 130 K=1,NG
         DETA(K)=0.D0
      DO 130 I=1,NC
         DS(K,I)=0.D0
         DO 120 M=1,NG
            IF(QT(M,I).EQ.0.D0) GOTO 120
            DS(K,I)=DS(K,I)-QT(M,I)*DLOG(TAU(M,K))*TAU(M,K)/T
 120     CONTINUE
         DETA(K)=DETA(K)+DS(K,I)*X(I)
 130  CONTINUE
C
      DO 150 I=1,NC
         TACT(I)=0.D0
         DO 140 K=1,NG
            H3(K,I)=(-S(K,I)*DETA(K)/ETA(K)+DS(K,I))/ETA(K)
            HH=H3(K,I)*(TETAR(K)-QT(K,I)*ETA(K)/S(K,I))
            TACT(I)=TACT(I)-HH
 140     CONTINUE
         TACT(I)=TACT(I)*ACT(I)
 150  CONTINUE
 160  CONTINUE
      RETURN
      END
C
C  *********************************************************************
C  *                                                                   *
C  *      PARMS CALCULATES SOME COMPOSITION-INDEPENDENT QUANTITIES:    *
C  *      "TAU", "S", AND "F", TO BE USED IN UNIMOD.  PARMS IS         *
C  *      CALLED PRIOR TO UNIMOD.                                      *
C  *                                                                   *
C  *********************************************************************
      SUBROUTINE PARMS(NC,NG,T)
C  *********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      COMMON /U/ RT(10,10),QT(10,10),TAU(10,10),S(10,10),F(10),Q(10),
     &           R(10),P(10,10)
C
      DO 10 I=1,NG
      DO 10 J=1,NG
         TAU(I,J)=DEXP(-P(I,J)/T)
  10  CONTINUE
C
      DO 20 I=1,NC
      DO 20 K=1,NG
         S(K,I)=0.D0
      DO 20 M=1,NG
         S(K,I)=S(K,I)+QT(M,I)*TAU(M,K)
  20  CONTINUE
C
      DO 30 I=1,NC
         F(I)=1.D0
      DO 30 J=1,NG
         F(I)=F(I)+QT(J,I)*DLOG(S(J,I))
  30  CONTINUE
      RETURN
      END
C
C***********************************************************************
C        Calculates adsorption onto activated carbon
C        using a generalized liquid-phase Polanyi isotherm.
C        Freundlich "K" (XK1) is in {(ug/gm) (L/ug)**(1/n)},
C        where:  QCAP = (XK1)*(CONC)**XNF
C***********************************************************************
      SUBROUTINE ADLIQ(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,VBM,KERR)
C***********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      COMMON /I/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
      COMMON /O/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
C
C        -- AQUEOUS SOLUBILITY BY "UNIFAC" --
C
      NPRNT=1
      KERR=0
      CALL AQSOL(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,TIE,VBM,NPRNT,KERR)
      IF(KERR.EQ.-1) GOTO 40
      DIFF=SOLUB-(CBULK/1000.D0)
      IF(DIFF.LE.0.D0) THEN
         KERR=-1
         WRITE(4,906)
         WRITE(7,906)
         WRITE(*,906)
         GOTO 40
      ENDIF
      IF(NVOL.NE.0) ORGDEN=FWT/VBM
      DENS=ORGDEN
C
C        >> POLANYI GENERALIZED ISOTHERM:
C           CZERO in {umol/L};  CS in {umol/L}
C
      CZERO=CBULK/FWT
      CS=(SOLUB*1000.D0)/FWT
      VOLM=FWT/DENS
      RHOM=(DENS*1.D06)/FWT
      RGAS=1.987D0
C
      NL=NL+1
      NJ=NL/2
      NL=2*NJ
      DP=DBLE(NL)
      OMAG=OMAG/2.D0
      CBEG=DLOG(CZERO/10.D0**OMAG)
      CEND=DLOG(CZERO*10.D0**OMAG)
      DIFF=CS-DEXP(CEND)
      IF(DIFF.LE.0.D0) THEN
         CEND=DLOG(0.99D0*CS)
         WRITE(4,907)
         WRITE(7,907)
         WRITE(*,907)
      ENDIF
      IF(CBEG.GE.CEND) THEN
         KERR=-1
         WRITE(4,908)
         WRITE(7,908)
         WRITE(*,908)
         GOTO 40
      ENDIF
      CINC=(CEND-CBEG)/DP
C
         CSAV=0.D0
         QSAV=0.D0
         SUMX=0.D0
         SUMY=0.D0
         SUMXX=0.D0
         SUMYY=0.D0
         SUMXY=0.D0
      DO 20 K=1,NL
         CLNC=CBEG+DBLE(K)*CINC
         CONC=DEXP(CLNC)
         ADSP=(RGAS*TT)*DLOG(CS/CONC)
         QCAP=(RHOM*W0)*DEXP(-BB*(ADSP/VOLM)**GM)
         QLNQ=DLOG(QCAP)
         SUMX=SUMX+CLNC
         SUMY=SUMY+QLNQ
         SUMXX=SUMXX+(CLNC)**2
         SUMYY=SUMYY+(QLNQ)**2
         SUMXY=SUMXY+(CLNC*QLNQ)
         IF(K.EQ.NJ) THEN
            CSAV=CONC*FWT
            QSAV=QCAP*FWT
         ENDIF
  20  CONTINUE
C
C        >> Frendlich "K" and "1/n" by linear regression:
C
      B0=(SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1=(DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF=B1
      XK2=DEXP(B0)
      XK1=(XK2*FWT)*(1.D0/FWT)**XNF
      RSQD=1.D0-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)
C
C        >> Calculate the Root-Mean-Square Error (RMSE):
C
         RMSE=0.D0
      DO 30 J=1,NL
         CLNC=CBEG+DBLE(J)*CINC
         CONC=DEXP(CLNC)
         ADSP=(RGAS*TT)*DLOG(CS/CONC)
         QCAP=(RHOM*W0)*DEXP(-BB*(ADSP/VOLM)**GM)
         QCAL=XK2*(CONC)**XNF
         RMSE=RMSE+((QCAL-QCAP)/QCAP)**2
  30  CONTINUE
         RMSE=DSQRT(RMSE/DP)*100.D0
         CBEG=DEXP(CBEG)*FWT
         CEND=DEXP(CEND)*FWT
  40  CONTINUE
      RETURN
C
C        -- FORMAT STATEMENTS --
C
 906  FORMAT(2X,'** ERROR: Bulk concentration exceeds solubility **')
 907  FORMAT(/,2X,
     & ' WARNING: Upper isotherm limit reset to 99% of C,sat.',/)
 908  FORMAT(2X,'** ERROR: Inappropriate isotherm regression limits **')
      END
C
C***********************************************************************
C        Calculates adsorption onto Calgon BPL activated carbon
C        using a generalized gas-phase Polanyi isotherm.
C        Freundlich "K" (XK1) is in {(ug/gm) (L/ug)**(1/n)},
C        where:  QCAP = (XK1)*(CONC)**XNF
C
C        THE EQUILIBRIUM ADSORBED CHEMICAL CONCENTRATION IS
C        ESTIMATED FOR A CHEMICAL-AIR-WATER MIXTURE --
C
C        RELHUM = RELATIVE HUMIDITY OF GAS STREAM, [60 - 100% R.H.]
C***********************************************************************
      SUBROUTINE ADVAP(RELHUM,IMOD,NVOL,IRNG,NC,TT,FWT,VBM,RNDX,KERR)
C***********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      COMMON /I/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
      COMMON /O/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
      COMMON /V/ NEQN,ANTA,ANTB,ANTC,ANTD,ANTE,TMIN,TMAX,PS,CS
C
C        >> Parameters in D-R equation:
C
C     W0=0.46D0
C     BB=3.37D-08
      NM=IDINT(GM)
C
      NSOL=1
      NPRNT=1
      KERR=-1
      CALL AQSOL(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,TIE,VBM,NPRNT,KERR)
      IF(KERR.EQ.-1) GOTO 40
C
      IF(RNDX.LT.1.D-03) THEN
         KERR=-1
         WRITE(4,914)
         WRITE(7,914)
         WRITE(*,914)
         GOTO 40
      ENDIF
C
         PS=DABS(ANTB)
      IF(ANTA.GT.1.D-03) THEN
         PS=DEXP(ANTA-ANTB/(TT+ANTC))
      ELSEIF(PS.LT.1.D-03) THEN
         KERR=-1
         WRITE(4,913)
         WRITE(7,913)
         WRITE(*,913)
         GOTO 40
      ENDIF
      PS=PS/760.D0
      IF(PS.GT.1.D0) THEN
         WRITE(4,912)
         WRITE(7,912)
         WRITE(*,912)
      ENDIF
      CS=(PS/0.08206D0/TT)*1.D06
      DIFF=(CS*FWT)-CBULK
      IF(DIFF.LE.0.D0) THEN
         KERR=-1
         WRITE(4,909)
         WRITE(7,909)
         WRITE(*,909)
         GOTO 40
      ENDIF
      IF(NVOL.NE.0) ORGDEN=FWT/VBM
      DENS=ORGDEN
      BETA=1.D0 - 0.34D0
      IF(RELHUM.LT.60.D0) THEN
         RELHUM=100.D0
         BETA=1.D0
      ENDIF
C
C        >> POLANYI GENERALIZED ISOTHERM;
C           CALCULATE REFRACTIVE INDEX "SCALE FACTORS"...
C
C           REFERENCE CHEMICAL : TOLUENE - [1]
C                         MW = 92.14 {G/GMOL}
C                    DENSITY = 0.8623 {G/CC}
C           REFRACTIVE INDEX = 1.4941
C
C           REFERENCE CHEMICAL : N-HEPTANE - [2]
C                         MW = 100.19 {G/GMOL}
C                    DENSITY = 0.6795 {G/CC}
C           REFRACTIVE INDEX = 1.3851
C
      REFMW1=92.14
      REFMW2=100.19
      REFDN1=0.8623
      REFDN2=0.6795
      RIREF1=1.4941
      RIREF2=1.3851
      OREF1=(REFMW1/REFDN1)*(RIREF1**2-1.)/(RIREF1**2+2.)
      OREF2=(REFMW2/REFDN2)*(RIREF2**2-1.)/(RIREF2**2+2.)
      SF1=(FWT/DENS)*(RNDX**2-1.)/(RNDX**2+2.)/OREF1
      SF2=(FWT/DENS)*(RNDX**2-1.)/(RNDX**2+2.)/OREF2
C
C        >> CZERO in {umol/L}; CBULK in {ug/L}
C
      CZERO=CBULK/FWT
      VOLM=FWT/DENS
      VH2O=18.015D0/1.D0
      RHOM=(DENS/FWT)*1.D06
      RGAS=1.987D0
C
      NL=NL+1
      NJ=NL/2
      NL=2*NJ
      DP=DBLE(NL)
      OMAG=OMAG/2.D0
      CBEG=DLOG(CZERO/10.D0**OMAG)
      CEND=DLOG(CZERO*10.D0**OMAG)
      DIFF=CS-DEXP(CEND)
      IF(DIFF.LE.0.D0) THEN
         CEND=DLOG(0.99D0*CS)
         WRITE(4,910)
         WRITE(7,910)
         WRITE(*,910)
      ENDIF
      IF(CBEG.GE.CEND) THEN
         KERR=-1
         WRITE(4,911)
         WRITE(7,911)
         WRITE(*,911)
         GOTO 40
      ENDIF
      CINC=(CEND-CBEG)/DP
C
         CSAV=0.D0
         QSAV=0.D0
         SUMX=0.D0
         SUMY=0.D0
         SUMXX=0.D0
         SUMYY=0.D0
         SUMXY=0.D0
      DO 20 K=1,NL
         CLNC=CBEG+DBLE(K)*CINC
         CONC=DEXP(CLNC)
         PI=(CONC*1.D-06)*(0.08206D0*TT)
      IF(IMOD.EQ.2) THEN
C
C        >> DUBININ-RADUSHKEVICH (D-R) EQUATION:
C
         ADSP=(RGAS*TT)*DLOG(PS/PI)
         QCAP=(RHOM*W0)*DEXP(-BB*(ADSP/SF1)**NM)
      ELSE
C
C        >> CALGON CHARACTERISTIC EQUATION (BPL):
C
         EL=DLOG10(PS/PI)
         YORD=(TT/BETA)*(EL/VOLM-DLOG10(100.D0/RELHUM)/VH2O)
         XORD=YORD/SF2
         QLOG = 1.71D0 - (1.46D-02*XORD)
     &                 - (1.65D-03*XORD**2)
     &                 - (4.11D-04*XORD**3)
     &                 + (3.14D-05*XORD**4)
     &                 - (6.75D-07*XORD**5)
         QQ=(10.D0**QLOG)/100.D0
         QCAP=(QQ*DENS/FWT)*1.D06
      ENDIF
         IF(QCAP.LE.1.D-03) THEN
            KERR=-1
            WRITE(4,915)
            WRITE(7,915)
            WRITE(*,915)
            GOTO 40
         ENDIF
         QLNQ=DLOG(QCAP)
         SUMX=SUMX+CLNC
         SUMY=SUMY+QLNQ
         SUMXX=SUMXX+(CLNC)**2
         SUMYY=SUMYY+(QLNQ)**2
         SUMXY=SUMXY+(CLNC*QLNQ)
         IF(K.EQ.NJ) THEN
            CSAV=CONC*FWT
            QSAV=QCAP*FWT
            RSAV=PI/PS
         ENDIF
  20  CONTINUE
C
C        >> Frendlich "K" and "1/n" by linear regression:
C
      B0=(SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1=(DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF=B1
      XK2=DEXP(B0)
      XK1=(XK2*FWT)*(1.D0/FWT)**XNF
      RSQD=1.D0-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)
C
C        >> Calculate the Root-Mean-Square Error (RMSE):
C
         RMSE=0.D0
      DO 30 J=1,NL
         CLNC=CBEG+DBLE(J)*CINC
         CONC=DEXP(CLNC)
         PI=(CONC*1.D-06)*(0.08206D0*TT)
      IF(IMOD.EQ.2) THEN
C
C        >> DUBININ-RADUSHKEVICH (D-R) EQUATION:
C
         ADSP=(RGAS*TT)*DLOG(PS/PI)
         QCAP=(RHOM*W0)*DEXP(-BB*(ADSP/SF1)**NM)
      ELSE
C
C        >> CALGON CHARACTERISTIC EQUATION (BPL):
C
         EL=DLOG10(PS/PI)
         YORD=(TT/BETA)*(EL/VOLM-DLOG10(100.D0/RELHUM)/VH2O)
         XORD=YORD/SF2
         QLOG = 1.71D0 - (1.46D-02*XORD)
     &                 - (1.65D-03*XORD**2)
     &                 - (4.11D-04*XORD**3)
     &                 + (3.14D-05*XORD**4)
     &                 - (6.75D-07*XORD**5)
         QQ=(10.D0**QLOG)/100.D0
         QCAP=(QQ*DENS/FWT)*1.D06
      ENDIF
         QCAL=XK2*(CONC)**XNF
         RMSE=RMSE+((QCAL-QCAP)/QCAP)**2
  30  CONTINUE
         RMSE=DSQRT(RMSE/DP)*100.D0
         CBEG=DEXP(CBEG)*FWT
         CEND=DEXP(CEND)*FWT
         CS=(CS*FWT)/1000.D0
C
C        >> WARN USER IF IN "PORE FILLING" REGIME...
C           i.e., if (Pi/Ps) > 0.2
C
      RLIM=0.2D0
      RATIO=PI/PS
      IF((RSAV.GE.RLIM).OR.(RATIO.GE.RLIM)) THEN
         WRITE(4,916)
         WRITE(7,916)
         WRITE(*,916)
      ENDIF
  40  CONTINUE
      RETURN
C
C        -- FORMAT STATEMENTS --
C
 909  FORMAT(2X,'** ERROR: Bulk concentration exceeds sat. pressure **')
 910  FORMAT(/,2X,
     & ' WARNING: Upper isotherm limit reset to 99% of C,sat.',/)
 911  FORMAT(2X,'** ERROR: Inappropriate isotherm regression limits **')
 912  FORMAT(/,2X,' WARNING: Saturation pressure exceeds 1 atm;',/,
     &       2X,'          chemical may be above its boiling point.',/)
 913  FORMAT(2X,'** ERROR: Vapor pressure information not supplied **')
 914  FORMAT(2X,'** ERROR: No value of refractive index was entered **')
 915  FORMAT(2X,'** ERROR: Polanyi correlation range is exceeded **')
 916  FORMAT(/,2X,' WARNING: (Pi/Ps) > 0.2, so the Polanyi results',/,
     &       2X,'          may correspond to capillary condensation.',/)
      END
C
C***********************************************************************
C
C   THIS PROGRAM USES THE D-R EQUATION TO PREDICT THE SURFACE LOADING
C   FOR A DESIRED GAS CONCENTRATION (ug/L).  THE PROGRAM THEN FINDS
C   THE FREUNDLICH "K" AND "1/N" PARAMETERS FOR WHICH THE D-R AND
C   FREUNDLICH SPREADING PRESSURES ARE EQUAL.
C
C   ALGORITHM DEVELOPED BY:  RANDY D. CORTRIGHT, GRADUATE STUDENT
C                            DAVID W. HAND, SENIOR RESEARCH ENGINEER
C                            [ June 1985 ]
C
C              MODIFIED BY:  TONY N. ROGERS, CHE GRADUATE STUDENT
C                            [ July 1991 ]
C
C   Freundlich "K" (XK1) is in {(ug/gm) (L/ug)**(1/n)},
C   where:  QA = (XK1)*(CONC)**XNF
C
C***********************************************************************
      SUBROUTINE SPEQ(NVOL,IRNG,NC,TT,FWT,VBM,RNDX,SPRD,BETA,XERR,KERR)
C***********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      COMMON /F/ TOL,IMAX
      COMMON /I/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
      COMMON /O/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
      COMMON /V/ NEQN,ANTA,ANTB,ANTC,ANTD,ANTE,TMIN,TMAX,PS,CS
C
C        >> Parameters in D-R equation:
C
C     W0=0.46D0
C     BB=3.37D-08
      NM=IDINT(GM)
C
      XERR=0.1
      SPRD=0.D0
      NSOL=1
      NPRNT=1
      KERR=-1
      CALL AQSOL(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,TIE,VBM,NPRNT,KERR)
      IF(KERR.EQ.-1) GOTO 40
      IF(NVOL.NE.0) ORGDEN=FWT/VBM
      DENS=ORGDEN
      RGAS=1.987D0
C
      IF(RNDX.LT.1.D-03) THEN
         KERR=-1
         WRITE(4,801)
         WRITE(7,801)
         WRITE(*,801)
         GOTO 40
      ENDIF
C
         PS=DABS(ANTB)
      IF(ANTA.GT.1.D-03) THEN
         PS=DEXP(ANTA-ANTB/(TT+ANTC))
      ELSEIF(PS.LT.1.D-03) THEN
         KERR=-1
         WRITE(4,802)
         WRITE(7,802)
         WRITE(*,802)
         GOTO 40
      ENDIF
      PS=PS/760.D0
      IF(PS.GT.1.D0) THEN
         WRITE(4,803)
         WRITE(7,803)
         WRITE(*,803)
      ENDIF
C
C        >> CS, CBULK in {ug/L}...
C
      CS=(PS/0.08206D0/TT)*(FWT*1.D06)
      DIFF=CS-CBULK
      IF(DIFF.LE.0.D0) THEN
         KERR=-1
         WRITE(4,804)
         WRITE(7,804)
         WRITE(*,804)
         GOTO 40
      ENDIF
C
C        >> DUBININ-RADUSKEVICH (D-R) CORRELATION;
C           CALCULATE REFRACTIVE INDEX "SCALE FACTOR"...
C
C           REFERENCE CHEMICAL : TOLUENE
C                         MW = 92.14 {G/GMOL}
C                    DENSITY = 0.8623 {G/CC}
C           REFRACTIVE INDEX = 1.4941
C
      REFMW=92.14
      REFDEN=0.8623
      RIREF=1.4941
      OREF=(REFMW/REFDEN)*(RIREF**2-1.)/(RIREF**2+2.)
      BETA=(FWT/DENS)*(RNDX**2-1.)/(RNDX**2+2.)/OREF
C
C        >> CALCULATE UPPER LIMIT OF SURFACE LOADING:
C
      QMIN=0.D0
      YMIN=0.D0
      PMAX=(CBULK*1.D-06/FWT)*(0.08206D0*TT)
      QMAX=W0*(1.D06*DENS/FWT)
     &    *DEXP(-BB*((RGAS*TT/BETA)*DLOG(PS/PMAX))**NM)
C
C        >> NUMERICAL INTEGRATION BY SIMPSON'S RULE:
C           (Step size is reduced until area converges)
C
         JMAX=NL+1
         JMAX=(JMAX/2)*2
         IF(JMAX.LE.0) JMAX=2
         SSAV=0.D0
   5  CONTINUE
         XSTEP=(QMAX-QMIN)/DBLE(JMAX)
         ICOUNT=0
         IFCN=0
         SUMY=0.D0
         QVAL=QMIN
         YVAL=YMIN
  10  CONTINUE
      IF(ICOUNT.LE.JMAX) THEN
      IF(ICOUNT.EQ.0) GOTO 15
C
C        >> CALCULATE THE VALUE OF d{ln C}/d{ln q}
C           AT THE GAS CONCENTRATION...
C           THIS IS EQUAL TO THE FREUNDLICH "N"
C
         QVAL=QVAL+XSTEP
C        IF(QVAL.LE.1.D-??) THEN
C           KERR=-1
C           WRITE(4,807)
C           WRITE(7,807)
C           WRITE(*,807)
C           GOTO 40
C        ENDIF
         ARG=QVAL/W0/(1.D06*DENS/FWT)
         VALUE=DLOG(ARG*1.D+25)-DLOG(1.D+25)
         PI=PS*DEXP(-DSQRT(VALUE/(-BB*(RGAS*TT/BETA)**NM)))
         YVAL=1.D0/((2.D0*BB)*DLOG(PS/PI)*(RGAS*TT/BETA)**NM)
C
  15  CONTINUE
         ICOUNT=ICOUNT+1
         ITST=((ICOUNT/2)*2)/ICOUNT
         ITST=2*(ITST+1)
         IFCN=IFCN+ITST
         WEIGHT=DBLE(ITST)
         SUMY=SUMY+(WEIGHT*YVAL)
         GOTO 10
      ENDIF
C
      YMAX=YVAL
      QMAX=QVAL
      IFCN=IFCN-2
      SUMY=SUMY-YMIN-YMAX
      SUMY=(SUMY/DBLE(IFCN))*(QMAX-QMIN)
      IF(DABS((SSAV-SUMY)/SUMY).GT.TOL) THEN
         JMAX=JMAX*2
         SSAV=SUMY
         GOTO 5
      ENDIF
      SPRD=SUMY
      IF(JMAX.LE.100000) THEN
         NL=JMAX
      ELSE
         NL=0
      ENDIF
C
C        >> CALCULATE THE FREUNDLICH PARAMETERS:
C
      XNF=QMAX/SPRD
      XK2=QMAX/(CBULK/FWT)**XNF
      XK1=(XK2*FWT)*(1.D0/FWT)**XNF
      CSAV=(PI/0.08206D0/TT)*(FWT*1.D06)
      QSAV=XK1*(CSAV)**XNF
C
C        >> LOWEST CONC. WHERE ERROR IS LESS THAN (XERR*100) PERCENT
C           BETWEEN THE D-R AND FREUNDLICH PREDICTIONS:
C
      CONC=CSAV
      CINC=CONC*TOL
      QTST=QSAV
      QINT=QMAX*FWT
  20  CONTINUE
      IF(DABS((QTST-QINT)/QINT).LE.XERR) THEN
         CONC=CONC-CINC
         PI=(CONC*1.D-06/FWT)*(0.08206D0*TT)
         QINT=W0*(1.D06*DENS)
     &       *DEXP(-BB*((RGAS*TT/BETA)*DLOG(PS/PI))**NM)
         QTST=XK1*(CONC)**XNF
         GOTO 20
      ENDIF
      CBEG=CONC
C
C        >> HIGHEST CONC. WHERE ERROR IS LESS THAN (XERR*100) PERCENT
C           BETWEEN THE D-R AND FREUNDLICH PREDICTIONS:
C
      CONC=CSAV
      CINC=CONC*TOL
      QTST=QSAV
      QINT=QMAX*FWT
  30  CONTINUE
      IF(DABS((QTST-QINT)/QINT).LE.XERR) THEN
         CONC=CONC+CINC
         PI=(CONC*1.D-06/FWT)*(0.08206D0*TT)
         QINT=W0*(1.D06*DENS)
     &       *DEXP(-BB*((RGAS*TT/BETA)*DLOG(PS/PI))**NM)
         QTST=XK1*(CONC)**XNF
         GOTO 30
      ENDIF
      CEND=CONC
C
C        >> CHECK VALIDITY OF ISOTHERM LIMITS:
C
      DIFF=CS-CEND
      CS=CS/1000.D0
      IF(DIFF.LE.0.D0) THEN
         KERR=-1
         WRITE(4,805)
         WRITE(7,805)
         WRITE(*,805)
         GOTO 40
      ENDIF
      IF(CBEG.GE.CEND) THEN
         KERR=-1
         WRITE(4,806)
         WRITE(7,806)
         WRITE(*,806)
      ENDIF
  40  CONTINUE
      RETURN
C
C        -- FORMAT STATEMENTS --
C
 801  FORMAT(2X,'** ERROR: No value of refractive index was entered **')
 802  FORMAT(2X,'** ERROR: Vapor pressure information not supplied **')
 803  FORMAT(/,2X,' WARNING: Saturation pressure exceeds 1 atm;',/,
     &       2X,'          chemical may be above its boiling point.',/)
 804  FORMAT(2X,'** ERROR: Bulk concentration exceeds sat. pressure **')
 805  FORMAT(2X,'** ERROR: Upper isotherm limit exceeds C,sat **')
 806  FORMAT(2X,'** ERROR: Inappropriate isotherm regression limits **')
 807  FORMAT(2X,'** ERROR: SPEQ integration step became too small **')
      END
C
C**********************************************************************
C Calculates adsorption onto activated carbon (liquid-phase) using
C the Uniform-Adsorbate Model developed by Manes and Hofer.
C Freundlich "K" (XK1) is in (ug/g)(L/ug)**1/n
C*********************************************************************
      SUBROUTINE HOFMAN(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,VBM,RNDX,KERR)
C*********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      COMMON /I/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
      COMMON /O/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
      COMMON /V/ NEQN,ANTA,ANTB,ANTC,ANTD,ANTE,TMIN,TMAX,PS,CS
C
C -- AQUEOUS SOLUBILITY BY UNIFAC --
C
      NPRNT=1
      KERR=0
      CALL AQSOL(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,TIE,VBM,NPRNT,KERR)
      IF(KERR.EQ.-1) GOTO 40
      DIFF=SOLUB-(CBULK/1000.D0)
      IF(DIFF.LE.0.D0) THEN
         KERR=-1
         WRITE(4,906)
         WRITE(7,906)
         WRITE(*,906)
         GOTO 40
      ENDIF
      IF(NVOL.NE.0) ORGDEN=FWT/VBM
      DENS=ORGDEN
C
C  -- CZERO in uM/L; CS in uM/L
C
      CZERO=CBULK/FWT
      CS=(SOLUB*1000.D0)/FWT
      VOLM=FWT/DENS
      RHOM=(DENS*1.D06)/FWT
      RGAS=1.987D0
C
      NL=NL+1
      NJ=NL/2
      NL=2*NJ
      DP=DBLE(NL)
      OMAG=OMAG/2.D0
      CBEG=DLOG(CZERO/10.D0**OMAG)
      CEND=DLOG(CZERO*10.D0**OMAG)
      DIFF=CS-DEXP(CEND)
      IF(DIFF.LE.0.D0) THEN
         CEND=DLOG(0.99D0*CS)
         WRITE(4,907)
         WRITE(7,907)
         WRITE(*,907)
      ENDIF
      IF(CBEG.GE.CEND) THEN
         KERR=-1
         WRITE(4,908)
         WRITE(7,908)
         WRITE(*,908)
         GOTO 40
      ENDIF
      CINC=(CEND-CBEG)/DP
C
      IF(RNDX.LT.1.D-03) THEN
         KERR=-1
         WRITE(4,914)
         WRITE(7,914)
         WRITE(*,914)
         GOTO 40
      ENDIF
C
      PS=DABS(ANTB)
      IF (ANTA.GT.1.D-03) THEN
          PS=DEXP(ANTA-ANTB/(ANTC+TT))
      ELSE IF (PS.LT.1.D-03) THEN
          KERR=-1
          WRITE(4,913)
          WRITE(7,913)
          WRITE(*,913)
          GOTO 40
      ENDIF
      PS=PS/760.D0
      IF(PS.GT.1.D0) THEN
         WRITE(4,912)
         WRITE(7,912)
         WRITE(*,912)
      ENDIF
      CSS=(PS/0.08206/TT)*1.D06
      DIFF=(CSS*FWT)-CBULK
      IF(DIFF.LE.0.D0) THEN
         KERR=-1
         WRITE(4,909)
         WRITE(7,909)
         WRITE(*,909)
         GOTO 40
      ENDIF
C
      CSAV=0.D0
      QSAV=0.D0
      SUMX=0.D0
      SUMY=0.D0
      SUMXX=0.D0
      SUMYY=0.D0
      SUMXY=0.D0
C
C  -- GAS PHASE D-R CORRELATION --
C
C     REFERENCE CHEMICAL: TOLUENE
C                   MW = 92.14 g/gmol
C              DENSITY = 0.8623 g/cc
C     REFRACTIVE INDEX = 1.4941
C
      REFMW=92.14D0
      REFDEN=0.8623D0
      REFVOL=REFMW/REFDEN
      RIREF=1.4941D0
      OREF=(RIREF**2-1.D0)/(RIREF**2+2.D0)
C
C  -- CORRELATING DIVISOR FOR WATER VAPOR ISOTHERM IS 0.28 --
C
      GAMMA=((RNDX**2-1.D0)/(RNDX**2+2.D0))/OREF - 0.28D0
C
      DO 10 K=1,NL
         CLNC=CBEG+DBLE(K)*CINC
         CONC=DEXP(CLNC)
         ADSP=(RGAS*TT)*DLOG(CS/CONC)
         ADSPRF=(ADSP/VOLM)*REFVOL/GAMMA
         VOLAD=W0*DEXP(-BB*(ADSPRF)**GM)
         QCAP=RHOM*VOLAD
         IF(QCAP.LE.1.D-03) THEN
            KERR=-1
            WRITE(4,915)
            WRITE(7,915)
            WRITE(*,915)
            GOTO 40
         ENDIF
         QLNQ=DLOG(QCAP)
         SUMX=SUMX+CLNC
         SUMY=SUMY+QLNQ
         SUMXX=SUMXX+(CLNC)**2
         SUMYY=SUMYY+(QLNQ)**2
         SUMXY=SUMXY+(CLNC*QLNQ)
         IF(K.EQ.NJ) THEN
            CSAV=CONC*FWT
            QSAV=QCAP*FWT
         ENDIF
  10  CONTINUE
C
C  -- FREUNDLICH "K" AND 1/n BY LINEAR REGRESSION --
C
      B0=(SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1=(DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF=B1
      XK2=DEXP(B0)
      XK1=(XK2*FWT)*(1.D0/FWT)**XNF
      RSQD=1.D0-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)
C
C  -- CALCULATE ROOT-MEAN-SQUARE ERROR (RMSE): --
C
      RMSE=0.D0
      DO 20 J=1,NL
         CLNC=CBEG+DBLE(K)*CINC
         CONC=DEXP(CLNC)
         ADSP=(RGAS*TT)*DLOG(CS/CONC)
         ADSPRF=(ADSP/VOLM)*REFVOL/GAMMA
         VOLAD=W0*DEXP(-BB*(ADSPRF)**GM)
         QCAP=RHOM*VOLAD
         QCAL=XK2*(CONC)**XNF
         RMSE=RMSE+((QCAL-QCAP)/QCAP)**2
  20  CONTINUE
      RMSE=DSQRT(RMSE/DP)**100.D0
      CBEG=DEXP(CBEG)*FWT
      CEND=DEXP(CEND)*FWT
  40  CONTINUE
      RETURN
C
C  -- FORMAT STATEMENTS --
C
 906  FORMAT(2X,'** ERROR: Bulk concentration exceeds solubility **')
 907  FORMAT(/,2X,
     & ' WARNING: Upper isotherm limit reset to 99% of C,sat.',/)
 908  FORMAT(2X,'** ERROR: Inappropriate isotherm regression limits **')
 914  FORMAT(2X,'** ERROR: No value of refractive index was entered **')
 913  FORMAT(2X,'** ERROR: Vapor pressure information not supplied **')
 912  FORMAT(/,2X,' WARNING: Saturation pressure exceeds 1 atm.',/)
 909  FORMAT(2X,'** ERROR: Bulk concentration exceeds sat. pressure **')
 915  FORMAT(2X,'** ERROR: D-R correlation range is exceeded **')
      END
C
C *********************************************************************
C       Calculates adsorption onto activated carbon (liquid-phase)
C       using the Non-Uniform Adsorbate Model developed by Hansen and
C       Fackler. This modification incorporates the variability
C       in composition of solute and solvent in the adsorbed phase.
C       Freundlich "K" (XK1) is in (ug/g)(L/ug)**1/n
C ********************************************************************
      SUBROUTINE HANFAC (NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,VBM,RNDX,KERR)
C ********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      COMMON /F/ TOL,IMAX
      COMMON /I/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
      COMMON /O/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
      COMMON /V/ NEQN,ANTA,ANTB,ANTC,ANTD,ANTE,TMIN,TMAX,PS,CS
      COMMON /INFO/ XMOLFR,XMOLFW,RGAS,VOLM,VH2O,ADSP,DADSP,DADSPW
C
C     -- AQUEOUS SOLUBILITY BY "UNIFAC" --
C
      NPRNT=1
      KERR=0
      CALL AQSOL(NSOL,NVOL,IRNG,NC,TT,FWT,SOLUB,TIE,VBM,NPRNT,KERR)
      IF (KERR .EQ. -1) GOTO 40
      DIFF = SOLUB-(CBULK/1000.D0)
      IF (DIFF .LE. 0.D0) THEN
         KERR = -1
         WRITE (4,906)
         WRITE (7,906)
         WRITE (*,906)
         GOTO 40
      ENDIF
      IF(NVOL.NE.0) ORGDEN=FWT/VBM
      DENS=ORGDEN
C
C        CZERO in uM/L; CS in uM/L
C
      CZERO=CBULK/FWT
      CS=(SOLUB*1000.D0)/FWT
      VOLM=FWT/DENS
      VH2O=18.015/1.0
      RHOM=(DENS*1.D06)/FWT
      RGAS=1.987D0
C
      NL=NL+1
      NJ=NL/2
      NL=2*NJ
      DP=DBLE(NL)
      OMAG=OMAG/2.D0
      CBEG=DLOG(CZERO/10.D0**OMAG)
      CEND=DLOG(CZERO*10.D0**OMAG)
      DIFF=CS-DEXP(CEND)
      IF (DIFF .LE. 0.D0) THEN
         CEND = DLOG(0.99D0*CS)
         WRITE(4,907)
         WRITE(7,907)
         WRITE(*,907)
      ENDIF
      IF (CBEG .GE. CEND) THEN
         KERR = -1
         WRITE(4,908)
         WRITE(7,908)
         WRITE(*,908)
         GOTO 40
      ENDIF
      CINC=(CEND-CBEG)/DP
C
      IF(RNDX.LT.1.D-03) THEN
         KERR=-1
         WRITE(4,914)
         WRITE(7,914)
         WRITE(*,914)
         GOTO 40
      ENDIF
C
      PS=DABS(ANTB)
      IF (ANTA.GT.1.D-03) THEN
          PS=DEXP(ANTA-ANTB/(ANTC+TT))
      ELSE IF (PS.LT.1.D-03) THEN
          KERR=-1
          WRITE(4,913)
          WRITE(7,913)
          WRITE(*,913)
          GOTO 40
      ENDIF
      PS=PS/760.D0
      IF(PS.GT.1.D0) THEN
         WRITE(4,912)
         WRITE(7,912)
         WRITE(*,912)
      ENDIF
      CSS=(PS/0.08206/TT)*1.D06
      DIFF=(CSS*FWT)-CBULK
      IF(DIFF.LE.0.D0) THEN
         KERR=-1
         WRITE(4,909)
         WRITE(7,909)
         WRITE(*,909)
         GOTO 40
      ENDIF
C
C  ***EQUATION FOR VAPOUR PRESSURE OF WATER OBTAINED FROM
C     THE PROPERTIES OF GASES AND LIQUIDS BY REID ET AL., 1987***
C
      X=1.D0-(TT/647.3)
      PH2O=221.2D0*DEXP((-7.76451*X+1.45838*X**1.5-2.77580*X**3-
     &                    1.23303*X**6)/(1.D0-X))
      PH2O=PH2O/1.013D0
C
      CSAV=0.D0
      QSAV=0.D0
      SUMX=0.D0
      SUMY=0.D0
      SUMXX=0.D0
      SUMYY=0.D0
      SUMXY=0.D0
C
C     ***GAS PHASE D-R CORRELATION***
C
C     ***REFERENCE CHEMICAL: TOLUENE
C                       MW = 92.14 g/gmol
C                  DENSITY = 0.8623 g/cc
C         REFRACTIVE INDEX = 1.4941
C
      REFMW=92.14D0
      REFDEN=0.8623D0
      REFVOL=REFMW/REFDEN
      RIREF=1.4941D0
      OREF=(RIREF**2-1.D0)/(RIREF**2+2.D0)
      GAMMA1=((RNDX**2-1.D0)/(RNDX**2+2.D0))/OREF - 0.28D0
      GAMMA2=((RNDX**2-1.D0)/(RNDX**2+2.D0))/OREF
      SF=(VOLM/REFVOL)*GAMMA2
      SFW=(VH2O/REFVOL)*0.28D0
C
C  ***HANSEN-FACKLER MODIFICATION***
C
      DO 10 K = 1,NL
         QCAP=0.D0
         CLNC=CBEG+DBLE(K)*CINC
         CONC=DEXP(CLNC)
         ADP=(RGAS*TT)*DLOG(CS/CONC)
         ADSPRF=(ADP/VOLM)*REFVOL/GAMMA1
         VOLAD=W0*DEXP(-BB*(ADSPRF)**GM)
C
C           WRITE( *,'(5X,A,1PE11.4)') 'VOLAD =',VOLAD
C           WRITE(12,'(5X,A,1PE11.4)') 'VOLAD =',VOLAD
C
         IF (VOLAD .LE. 1.D-70) THEN
             VOLAD=1.D-70
         ENDIF
         QC=VOLAD*RHOM
         ZVOLAD=DLOG(VOLAD)
         ADSP=(ADSPRF/REFVOL)*GAMMA2*VOLM
C        PRINT *,'ADSP= ',ADSP
         ADSPW=(ADSPRF/REFVOL)*0.28D0*VH2O
         PI=PS*DEXP(-ADSP/(RGAS*TT))
         CON=(PI/0.08206/TT)*1.D06
         PIW=PH2O*DEXP(-ADSPW/(RGAS*TT))
         CONW=(PIW/0.08206/TT)*1.D06
         XMOLFR=CON/(CON+CONW)
         IF (XMOLFR .LE. 1.D-70) THEN
             XMOLFR=1.D-70
         ENDIF
         XMOLFW=CONW/(CON+CONW)
         DO 20 J =  K,1,-1
            ZDVOL=ZVOLAD*DBLE(K)
            XUMVOL=ZVOLAD*DBLE(J)
            SUMVOL=DEXP(XUMVOL)
            IF (SUMVOL .LE. 1.D-70) THEN
                SUMVOL=1.D-70
            ENDIF
            IF (J .EQ. K) THEN
                DELVOL=DEXP(ZDVOL)
              ELSE
                DELVOL=SUMVOL-DELVOL
            ENDIF
            DADSP=SF*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
C           PRINT *,'DADSP= ',DADSP
            DADSPW=SFW*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
C
            A=1.0D-70
            CALL GOLDEN(A,IMAX,TOL,N,X,FX,IERR)
            IF (IERR .EQ. -1) GOTO 40
            XADS=X
            XADSW=1.D0-X
C
            DADS=(XADS*DELVOL)/(XADS*VOLM*1.D-6+XADSW*VH2O*1.D-6) -
     +           (XMOLFR*DELVOL)/(XMOLFR*VOLM*1.D-6+XMOLFW*VH2O*1.D-6)
            QCAP=QCAP+DADS
            DELVOL=SUMVOL
 20      CONTINUE
         QCAP=QCAP+QC
C        PRINT *,'QCAP= ',QCAP
         IF(QCAP.LE.1.D-03) THEN
            KERR=-1
            WRITE(4,915)
            WRITE(7,915)
            WRITE(*,915)
            GOTO 40
         ENDIF
         QLNQ=DLOG(QCAP)
         SUMX=SUMX+CLNC
         SUMY=SUMY+QLNQ
         SUMXX=SUMXX+(CLNC)**2
         SUMYY=SUMYY+(QLNQ)**2
         SUMXY=SUMXY+(CLNC*QLNQ)
         IF (K .EQ. NJ) THEN
            CSAV=CONC*FWT
            QSAV=QCAP*FWT
         ENDIF
 10   CONTINUE
C
C     ---Freundlich "K" and "1/n" by linear regression---
C
      B0=(SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1=(DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF=B1
      XK2=DEXP(B0)
      XK1=(XK2*FWT)*(1.D0/FWT)**XNF
      RSQD=1.D0-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)
C
C     ***Calculate the Root-Mean-Square Error (RMSE)***
C
         RMSE=0.D0
         K=0
         J=0
      DO 30 K=1,NL
         QCAP=0.D0
         CLNC=CBEG+DBLE(K)*CINC
         CONC=DEXP(CLNC)
         ADP=(RGAS*TT)*DLOG(CS/CONC)
         ADSPRF=(ADP/VOLM)*REFVOL/GAMMA1
         VOLAD=W0*DEXP(-BB*(ADSPRF)**GM)
C
C           WRITE( *,'(5X,A,1PE11.4)') 'VOLAD =',VOLAD
C           WRITE(12,'(5X,A,1PE11.4)') 'VOLAD =',VOLAD
C
         IF (VOLAD .LE. 1.D-70) THEN
             VOLAD=1.D-70
         ENDIF
         QC=VOLAD*RHOM
         ZVOLAD=DLOG(VOLAD)
         ADSP=(ADSPRF/REFVOL)*GAMMA2*VOLM
         ADSPW=(ADSPRF/REFVOL)*0.28D0*VH2O
         PI=PS*DEXP(-ADSP/(RGAS*TT))
         CON=(PI/0.08206/TT)*1.D06
         PIW=PH2O*DEXP(-ADSPW/(RGAS*TT))
         CONW=(PIW/0.08206/TT)*1.D06
         XMOLFR=CON/(CON+CONW)
         IF (XMOLFR .LE. 1.D-70) THEN
             XMOLFR=1.D-70
         ENDIF
         XMOLFW=CONW/(CON+CONW)
         DO 35 J =  K,1,-1
            ZDVOL=ZVOLAD*DBLE(K)
            XUMVOL=ZVOLAD*DBLE(J)
            SUMVOL=DEXP(XUMVOL)
            IF (SUMVOL .LE. 1.D-70) THEN
                SUMVOL=1.D-70
            ENDIF
            IF (J .EQ. K) THEN
                DELVOL=DEXP(ZDVOL)
              ELSE
                DELVOL=SUMVOL-DELVOL
            ENDIF
            DADSP=SF*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
            DADSPW=SFW*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
C
            A=1.0D-70
            CALL GOLDEN(A,IMAX,TOL,N,X,FX,IERR)
            IF (IERR .EQ. -1) GOTO 40
            XADS=X
            XADSW=1.D0-X
C
            DADS=(XADS*DELVOL)/(XADS*VOLM*1.D-6+XADSW*VH2O*1.D-6) -
     +           (XMOLFR*DELVOL)/(XMOLFR*VOLM*1.D-6+XMOLFW*VH2O*1.D-6)
            QCAP=QCAP+DADS
            DELVOL=SUMVOL
 35      CONTINUE
         QCAP=QCAP+QC
         IF (QCAP .LE. 1.D-70) GOTO 30
         QCAL=XK2*(CONC)**XNF
         RMSE=RMSE+((QCAL-QCAP)/QCAP)**2
 30   CONTINUE
         RMSE=DSQRT(RMSE/DP)*100.D0
         CBEG=DEXP(CBEG)*FWT
         CEND=DEXP(CEND)*FWT
 40   CONTINUE
      RETURN
C
C     ---FORMAT STATEMENTS---
C
 906  FORMAT(2X,'** ERROR: Bulk concentration exceeds solubility **')
 907  FORMAT(/,2X,
     & ' WARNING: Upper isotherm limit reset to 99% of C,sat.',/)
 908  FORMAT(2X,'** ERROR: Inappropriate isotherm regression limits **')
 914  FORMAT(2X,'** ERROR: No value of refractive index was entered **')
 913  FORMAT(2X,'** ERROR: Vapor pressure information not supplied **')
 912  FORMAT(/,2X,' WARNING: Saturation pressure exceeds 1 atm.',/)
 909  FORMAT(2X,'** ERROR: Bulk concentration exceeds sat.pressure **')
 915  FORMAT(2X,'** ERROR: D-R correlation range is exceeded **')
      END
C
C *********************************************************************
C              GOLDEN - SECTION SUBROUTINE
C              ---------------------------
C      USED TO SOLVE THE NON-LINEAR EQUATION GENERATED
C      BY THE SUBROUTINE 'HANFAC'
C *********************************************************************
      SUBROUTINE GOLDEN (A, IMAX, TOL, N, X, FX, IERR)
C *********************************************************************
      IMPLICIT REAL*8 (A-H,O-Z)
      COMMON /L/ TT,NG
      COMMON /INFO/ XMOLFR,XMOLFW,RGAS,VOLM,VH2O,ADSP,DADSP,DADSPW
C
C     ***ENTER OBJECTIVE FUNCTION TO BE MINIMIZED***
C
C     OBJECT(X)=(-DADSP/VOLM + (RGAS*TT/VOLM)*DLOG(X/XMOLFR) +
C    &           DADSPW/VH2O - (RGAS*TT/VH2O)*DLOG((1.D0-X)/XMOLFW))**2
      OBJECT(X)=DABS(-DADSP/VOLM + (RGAS*TT/VOLM)*DLOG(X/XMOLFR) +
     &               DADSPW/VH2O - (RGAS*TT/VH2O)*DLOG((1.D0-X)/XMOLFW))
C
C     ***STATEMENT FUNCTION TO IMPLEMENT GOLDEN SECTION***
C
      SECT(X,Y) = X + 0.618D0*Y
C
C
      KFLAG=0
      N=0
      F1=OBJECT(A)
      FSAVE=F1
      B=0.999999D0
      F2=OBJECT(B)
      IF (F2 .GT. FSAVE) GOTO 10
      FSAVE=F2
 10   UNC=B-A
      IF (UNC .LE. TOL) GOTO 45
      IF (N .EQ. IMAX) GOTO 999
      IF (N .EQ. 0) GOTO 15
      IF (KFLAG .EQ. 1) GOTO 30
      IF (KFLAG .EQ. 2) GOTO 40
C
 15   X1=SECT(B,-UNC)
      IF (X1 .GE. 1.D0) THEN
          X1=0.99999D0
      ENDIF
      FX1=OBJECT(X1)
      IF (N .GT. 0) GOTO 25
C
 20   X2=SECT(A,UNC)
      IF (X2 .GE. 1.D0) THEN
          X2=0.99999D0
      ENDIF
      FX2=OBJECT(X2)
C
 25   N=N+1
      IF (FX1 .GT. FX2) GOTO 35
C
C     ***BRANCH FOR F(X1) < F(X2)***
C
      KFLAG=1
      B=X2
        GOTO 10
 30   X2=X1
      FX2=FX1
        GOTO 15
C
C     ***BRANCH FOR F(X1) > F(X2)***
C
 35   KFLAG=2
      A=X1
        GOTO 10
 40   X1=X2
      FX1=FX2
        GOTO 20
C
 45   X=(A+B)/2.D0
      IF (X .GT. 1.D0) GOTO 998
C     PRINT *,'X= ',X
      FX=OBJECT(X)
C     PRINT *,'FX= ',FX
      IERR=1
      RETURN
C
C     ***NON-CONVERGENCE FAILURE MESSAGE***
C
 998  IERR=-1
      WRITE (4,201)
      WRITE (7,201)
      WRITE (*,201)
 201  FORMAT(//,1X,'******  ERROR :  ROOT IS NOT BOUNDED',/)
      GOTO 1000
 999  IERR=-1
      WRITE (4,200) N
      WRITE (7,200) N
      WRITE (*,200) N
 200  FORMAT(//,1X,'******  ERROR :  SUBROUTINE GOLDEN DID NOT FIND THE
     &ROOT AFTER ',I4,'  ITERATIONS',/,//)
 1000 END
C
C***********************************************************************
C
C             << UNIFAC BINARY LIQUID-LIQUID FLASH ROUTINE >>
C
C     NEWTON-RAPHSON ALGORITHM (GAUSS-JORDAN MAXIMUM PIVOT STRATEGY)
C-----------------------------------------------------------------------
C   THIS IS A SUBROUTINE TO IMPLEMENT THE NEWTON-RAPHSON ALGORITHM FOR
C   SOLVING SYSTEMS OF NONLINEAR ALGEBRAIC EQUATIONS.  A VARIATION OF
C   THE GAUSS-JORDAN MAXIMUM PIVOT STRATEGY IS EMPLOYED TO DETERMINE
C   THE INVERSE OF THE JACOBIAN MATRIX.  THE CORRECTION FACTORS ARE
C   CALCULATED IN AN ITERATIVE MANNER TO BRING THE ADJUSTABLE VARIABLES
C   WITHIN A SPECIFIED TOLERANCE.
C
C                  MM = NUMBER OF COLUMNS IN MATRIX C
C                  NN = NUMBER OF ROWS IN MATRIX C
C
C***********************************************************************
      SUBROUTINE NEWTON(NN, MAXIT, TOL, XGUESS, XX, FF, NPRNT, IERR)
C***********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      DIMENSION  CC(2,3),IR(2),DX(2),XOLD(2),XX(2),
     &           FOLD(2),FF(2),XGUESS(2),X1(10),X2(10),
     &           ACT1(10),DACT1(10,10),TACT1(10),DLACT1(10),
     &           ACT2(10),DACT2(10,10),TACT2(10),DLACT2(10)
      COMMON /L/ TT,NG
C
C        -- SET WIDTH OF JACOBIAN MATRIX --
C
         MM = NN + 1
C
C        -- INITIALIZATION (CC, DX, XX, XOLD, FOLD) --
C
         IERR = 0
      DO 30 L=1,NN
         XX(L) = XGUESS(L)
         XOLD(L) = XGUESS(L)
         FF(L) = 0.D0
         FOLD(L) = 0.D0
         DX(L) = 0.D0
            DO 20 LL=1,MM
               CC(L,LL) = 0.D0
  20        CONTINUE
  30  CONTINUE
C
      NDIF=1
      NACT=0
      CALL PARMS(NN,NG,TT)
C
C        << START OF NEWTON-RAPHSON ITERATION LOOP >>
C
         IMAX = MAXIT + 1
      DO 180 ITER=1,IMAX
         INDEX = ITER - 1
C
C        -- EVALUATE FUNCTION VECTOR (FF) --
C
         X1(1) = XX(1)
         X1(2) = 1.D0 - X1(1)
         X2(1) = XX(2)
         X2(2) = 1.D0 - X2(1)
C
         CALL UNIMOD(NDIF,NACT,NN,NG,TT,X1,ACT1,DACT1,TACT1)
         CALL UNIMOD(NDIF,NACT,NN,NG,TT,X2,ACT2,DACT2,TACT2)
C
C        -- CALCULATE ROOT-MEAN-SQUARE ERROR (RMSE) --
C
            LOGIC = 0
            RMSE1 = 0.D0
            RMSE2 = 0.D0
            RMSE3 = 0.D0
         DO 40 J=1,NN
            FF(J) = X1(J)*ACT1(J) - X2(J)*ACT2(J)
            IF(DABS(FF(J)).GE.TOL) LOGIC = -1
            DIF1 = FF(J) - FOLD(J)
            DIF2 = XX(J) - XOLD(J)
            RMSE1 = RMSE1 + FF(J)**2
            RMSE2 = RMSE2 + (DIF1)**2
            RMSE3 = RMSE3 + (DIF2)**2
  40     CONTINUE
            IF(LOGIC.EQ.INDEX) GOTO 200
            RMSE1 = DSQRT(RMSE1/DBLE(NN))
            RMSE2 = DSQRT(RMSE2/DBLE(NN))
            RMSE3 = DSQRT(RMSE3/DBLE(NN))
C
C        -- TEST FOR CONVERGENCE OF SOLUTION --
C
            IF(RMSE1.GE.TOL) GOTO 50
            IF(RMSE2.GE.TOL) GOTO 50
            IF(RMSE3.GE.TOL) GOTO 50
            IF(LOGIC.EQ.-1)  GOTO 50
         GOTO 200
  50     CONTINUE
C
C        -- SAVE PREVIOUS ITERATION --
C
         DO 60 I=1,NN
            FOLD(I) = FF(I)
            XOLD(I) = XX(I)
  60     CONTINUE
C
C        -- LOAD PARTIAL DERIVATIVES IN JACOBIAN --
C
         CC(1,1) =  ACT1(1) + XX(1)*DACT1(1,1)
         CC(1,2) = -ACT2(1) - XX(2)*DACT2(1,1)
         CC(2,1) = -ACT1(2) + (1.D0 - XX(1))*DACT1(2,1)
         CC(2,2) =  ACT2(2) - (1.D0 - XX(2))*DACT2(2,1)
C
C        -- FINISH LOADING "CC" MATRIX WITH "FF" VECTOR --
C
         DO 70 I=1,NN
            CC(I,MM) = -FF(I)
  70     CONTINUE
C
C        ** GAUSS-JORDAN ALGORITHM **
C
C        -- INITIALIZE ALL VECTORS AND MATRICES --
C
         DO 80 I=1,NN
            DX(I) = 0.D0
            IR(I) = 0
            JJ = 0
            JM = 0
  80     CONTINUE
C
         DO 140 K=1,NN
            PK = 0.D0
C
C        -- LOCATE PIVOT ELEMENT --
C
            DO 100 I=1,NN
               IF(I.EQ.IR(I)) GOTO 100
               DO 90 IK=1,NN
                  PP = DABS(CC(I,IK))
                  IF(PP.LT.PK) GOTO 90
                  PK = PP
                  JJ = I
                  JM = IK
  90           CONTINUE
 100        CONTINUE
               IR(JJ) = JJ
C
C        -- NORMALIZATION STEP --
C
            DO 110 JR=1,MM
               IF(JM.EQ.JR) GOTO 110
               IF(DABS(CC(JJ,JM)).LE.1.D-25) GOTO 195
               CC(JJ,JR) = CC(JJ,JR) / CC(JJ,JM)
 110        CONTINUE
               CC(JJ,JM) = 1.D0
C
C        -- REDUCTION STEP --
C
            DO 130 I=1,NN
               IF(I.EQ.JJ) GOTO 130
               DO 120 JR=1,MM
                  IF(JR.EQ.JM) GOTO 120
                  CC(I,JR) = CC(I,JR) - CC(I,JM) * CC(JJ,JR)
 120           CONTINUE
               CC(I,JM) = 0.D0
 130        CONTINUE
 140     CONTINUE
C
C        ** END OF GAUSS-JORDAN MAXIMUM PIVOT ROUTINE **
C
C        -- RECOVER THE SOLUTION VECTOR --
C
         DO 160 I=1,NN
            DO 150 J=1,NN
               IF((CC(I,J).LT.1.D0) .OR.
     &            (CC(I,J).GT.1.D0)) GOTO 150
               DX(J) = CC(I,MM)
 150        CONTINUE
 160     CONTINUE
C
C        -- CORRECT ELEMENTS OF THE "XX" VECTOR --
C
         DO 170 I=1,NN
            XX(I) = XX(I) + DX(I)
            IF(XX(I).LT.0.D0) XX(I)=0.D0
            IF(XX(I).GT.1.D0) XX(I)=1.D0
 170     CONTINUE
 180  CONTINUE
C
C        << END OF ITERATION LOOP >>
C
C        -- CONVERGENCE FAILURE MESSAGE (MAXIT REACHED) --
C
 190  CONTINUE
      IERR = -1
      IF(NPRNT.NE.0) THEN
         WRITE(4,901) INDEX
         WRITE(7,901) INDEX
         WRITE(*,901) INDEX
      ENDIF
      RETURN
C
C        -- "DIVISION UNDERFLOW" ERROR MESSAGE --
C
 195  CONTINUE
      IERR = -1
      IF(NPRNT.NE.0) THEN
         WRITE(4,902)
         WRITE(7,902)
         WRITE(*,902)
      ENDIF
      RETURN
C
C        -- RETURN SOLUTION --
C
 200  CONTINUE
      IERR = 1
      RETURN
C
C        -- FORMAT STATEMENTS --
C
 901  FORMAT(5X,
     & 'Solubility algorithm did not converge after ',
     & I3,' iterations.',/,
     & 5X,'Results from the last iteration are discarded by NEWTON.',/)
 902  FORMAT(2X,'** ERROR: Division underflow in subroutine NEWTON **')
      END
C
C***********************************************************************
C
C        >> Calculates aqueous solubility using UNIFAC
C
C***********************************************************************
      SUBROUTINE AQSOL(NSOL,NVOL,IRNG,NC,TEMP,FWT,SOLUB,TIE,VBM,
     &                 NPRNT,KERR)
C***********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      DIMENSION  XGUESS(2),XSOLN(2),FF(2),XMW(10),
     &           X1(10),X2(10),XE(2),IE(2),
     &           ACT1(10),DACT1(10,10),TACT1(10),DLACT1(10),
     &           ACT2(10),DACT2(10,10),TACT2(10),DLACT2(10)
      COMMON /F/ TOL,IMAX
      COMMON /L/ TT,NG
C
      KSAV=KERR
      KERR=0
      TT=TEMP
      TIE=0.D0
C
         MODEL=0
      CALL FGRP(MODEL,NSOL,NVOL,IRNG,NC,NG,XMW,FWT,VBM,NPRNT,JERR)
         IF(JERR.EQ.-1) GOTO 30
         IF(NSOL.NE.0) GOTO 25
         XGUESS(1)=1.D0
         XGUESS(2)=0.D0
      CALL NEWTON(NC,IMAX,TOL,XGUESS,XSOLN,FF,NPRNT,IERR)
         IF(IERR.EQ.-1) GOTO 25
C
C        -- SORT COMPOSITION (DESCENDING ORDER) --
C
      DO 10 I=1,NC
         IF(XSOLN(I).LE.0.D0) GOTO 25
         IE(I)=1
      DO 10 J=1,NC
         IF(J.EQ.I) GOTO 10
         DIFF=DABS((XSOLN(J)-XSOLN(I))/XSOLN(I))*100.D0
         IF(DIFF.LE.0.1D0) GOTO 25
         IF(XSOLN(I).LT.XSOLN(J)) IE(I)=IE(I)+1
  10  CONTINUE
      DO 20 I=1,NC
         XE(I) = XSOLN(IE(I))
  20  CONTINUE
         X1(1) = XE(1)
         X1(2) = 1.D0 - X1(1)
         X2(1) = XE(2)
         X2(2) = 1.D0 - X2(1)
C
C        -- CONVERT MOLE FRACTION TO "PPMW" --
C
      XMF = 1.D0 - XE(1)
      XE(1) = 1.D06/(1.D0+((1.D0/XE(2))-1.D0)*XMW(2)/XMW(1))
      XE(2) = 1.D06/(1.D0+((1.D0/XMF)-1.D0)*XMW(1)/XMW(2))
      SOLUB = XE(2)
      TIE = XE(1)
C
  25  CONTINUE
      KERR=1
      IF((SOLUB.LE.TOL).AND.(KSAV.EQ.0)) THEN
         KERR=-1
         IF(NPRNT.NE.0) THEN
            WRITE(4,905)
            WRITE(7,905)
            WRITE(*,905)
         ENDIF
      ENDIF
      GOTO 40
  30  CONTINUE
      KERR=-1
  40  CONTINUE
      RETURN
C
C        -- FORMAT STATEMENTS --
C
 905  FORMAT(2X,'** ERROR: Problem with value of solubility **')
      END
C
C***********************************************************************
C
C        >> Calculates dimensionless Henry's Constants using UNIFAC
C
C***********************************************************************
      SUBROUTINE HENRY(TEMP,HLC,GAMMA,PVAP,NPRNT,MERR)
C***********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      DIMENSION  XX(10),ACT(10),DACT(10,10),TACT(10),DLACT(10),XMW(10)
      COMMON /V/ NEQN,ANTA,ANTB,ANTC,ANTD,ANTE,TMIN,TMAX,PS,CS
C
      MERR=0
      GAMMA=0.D0
      HLC=0.D0
      PRESS=1.D0
      FWT=0.D0
C
      CALL FGRP(0,0,1,0,2,NG,XMW,FWT,1.D0,0,JERR)
      IF(JERR.EQ.-1) THEN
         MERR=-1
         GOTO 5
      ENDIF
         XX(1)=1.D0
         XX(2)=0.D0
      CALL PARMS(2,NG,TEMP)
      CALL UNIMOD(0,0,2,NG,TEMP,XX,ACT,DACT,TACT)
      GAMMA = ACT(2)
C
C        >> VAPOR PRESSURE -- DIPPR[=101], OR ANTOINE[=0]:
C
   5  CONTINUE
      PVAP=DABS(ANTB)
      IF(ANTA.GT.1.D-03) THEN
         IF(NEQN.EQ.101) THEN
            IF((TEMP.LE.TMIN).OR.(TEMP.GE.TMAX)) THEN
               IF(NPRNT.NE.0) WRITE(11,903)
               ANTA=0.D0
               ANTB=0.D0
               MERR=-1
               GOTO 10
            ENDIF
            PVAP=DEXP(ANTA+(ANTB/TEMP)+(ANTC*DLOG(TEMP))
     &          +(ANTD*(TEMP**ANTE)))*(760.D0/1.01325D+05)
         ELSEIF(NEQN.EQ.0) THEN
            TC=TEMP-273.15D0
            IF((TC.LE.TMIN).OR.(TC.GE.TMAX)) THEN
               IF(NPRNT.NE.0) WRITE(11,903)
               ANTA=0.D0
               ANTB=0.D0
               MERR=-1
               GOTO 10
            ENDIF
            PVAP=DEXP(ANTA-ANTB/((TC+273.15D0)+ANTC))
         ENDIF
      ELSEIF(PVAP.LT.1.D-03) THEN
         IF(NPRNT.NE.0) WRITE(11,901)
         MERR=-1
         GOTO 10
      ENDIF
      ANTA=0.D0
      ANTB=PVAP
      PSAT=PVAP/760.D0
      IF(PSAT.GT.PRESS) THEN
         IF(NPRNT.NE.0) WRITE(11,902)
      ENDIF
      IF(MERR.EQ.-1) GOTO 10
C
C        >> HENRY'S LAW CONSTANT FOR COMPONENT [2]:
C
      HATM = GAMMA*PSAT
      HLC = HATM*(XMW(1)/1000.D0)/(0.08206D0*TEMP)
C
C        >> SET VALUE OF "MERR" (ERROR FLAG):
C
      MERR=1
  10  CONTINUE
      RETURN
C
C        -- FORMAT STATEMENTS --
C
 901  FORMAT(2X,'** ERROR: No vapor pressure for next chemical **')
 902  FORMAT(/,2X,' WARNING: Saturation pressure exceeds 1 atm;',/,
     &       2X,'          chemical may be above its boiling point.',/)
 903  FORMAT(2X,'** ERROR: Temp. is outside v.p. correlation limit **')
      END
C
C***********************************************************************
C
C        >> Calculates octanol/water partitioning using UNIFAC
C
C***********************************************************************
      SUBROUTINE PARTC(TEMP,OCTDEN,WATDEN,XKOW,XLGK,MERR)
C***********************************************************************
      IMPLICIT REAL*8(A-H,O-Z)
      DIMENSION  XGUESS(2),XSOLN(2),FF(2),XMW(10),
     &           X1(10),X2(10),XE(2),IE(2),MI(10,2),
     &           ACT1(10),DACT1(10,10),TACT1(10),DLACT1(10),
     &           ACT2(10),DACT2(10,10),TACT2(10),DLACT2(10)
      COMMON /F/ TOL,IMAX
      COMMON /G/ MS(10,10,2),NMAX
      COMMON /L/ TT,NG
C
C        >> INITIALIZE:
C
         MERR=0
         XKOW=0.D0
         XLGK=0.D0
         TT=TEMP
         NI=NMAX
      DO 10 J=1,10
      DO 10 K=1,2
         MI(J,K)=MS(2,J,K)
      DO 10 I=1,3
         MS(I,J,K)=0
  10  CONTINUE
C
C        >> LOAD UNIFAC GROUPS FOR [1]-WATER AND [2]-OCTANOL:
C
      NMAX=3
      MS(1,1,1)=17
      MS(1,1,2)=1
C
      MS(2,1,1)=1
      MS(2,1,2)=1
      MS(2,2,1)=2
      MS(2,2,2)=7
      MS(2,3,1)=15
      MS(2,3,2)=1
C
C        >> OCTANOL/WATER EQUILIBRIUM:
C
      NC=2
      NPRNT=0
      FWT=0.D0
      CALL FGRP(0,0,1,0,NC,NG,XMW,FWT,1.D0,NPRNT,JERR)
      IF(JERR.EQ.-1) GOTO 40
         XGUESS(1)=1.D0
         XGUESS(2)=0.D0
      CALL NEWTON(NC,IMAX,TOL,XGUESS,XSOLN,FF,NPRNT,IERR)
      IF(IERR.EQ.-1) GOTO 40
C
C        >> SORT COMPOSITION (DESCENDING ORDER):
C
      DO 20 I=1,NC
         IF(XSOLN(I).LE.0.D0) GOTO 40
         IE(I)=1
      DO 20 J=1,NC
         IF(J.EQ.I) GOTO 20
         DIFF=DABS((XSOLN(J)-XSOLN(I))/XSOLN(I))*100.D0
         IF(DIFF.LE.0.1D0) GOTO 40
         IF(XSOLN(I).LT.XSOLN(J)) IE(I)=IE(I)+1
  20  CONTINUE
      DO 30 I=1,NC
         XE(I) = XSOLN(IE(I))
  30  CONTINUE
C
C        >> MOLE FRACTIONS (INFINITE DILUTION OF CHEMICAL [3]):
C
      X1(1) = XE(1)
      X1(2) = 1.D0 - X1(1)
      X1(3) = 0.D0
      X2(1) = XE(2)
      X2(2) = 1.D0 - X2(1)
      X2(3) = 0.D0
C
C        >> PARTITIONING FOR DISTRIBUTED CHEMICAL:
C
      NC=3
      IF(NI.GT.NMAX) NMAX=NI
      DO 35 J=1,10
      DO 35 K=1,2
         MS(NC,J,K)=MI(J,K)
  35  CONTINUE
      FWT=0.D0
      CALL FGRP(0,0,1,0,NC,NG,XMW,FWT,1.D0,NPRNT,JERR)
      IF(JERR.EQ.-1) GOTO 40
      CALL PARMS(NC,NG,TEMP)
      CALL UNIMOD(0,0,NC,NG,TEMP,X1,ACT1,DACT1,TACT1)
      CALL UNIMOD(0,0,NC,NG,TEMP,X2,ACT2,DACT2,TACT2)
C
C        >> XKOW = PARTITION COEFFICIENT
C           XLGK = BASE-10 LOGARITHM OF XKOW
C
      PHASEW = 1.D0 / (X1(1)/WATDEN + X1(2)/OCTDEN)
      PHASEO = 1.D0 / (X2(1)/WATDEN + X2(2)/OCTDEN)
      RATIO = PHASEO/PHASEW
C     WRITE(*,*) 'RATIO = ',RATIO
      XKOW = (PHASEO/PHASEW)*(ACT1(3)/ACT2(3))
      XLGK = DLOG10(XKOW)
C
C        >> SET VALUE OF "MERR" (ERROR FLAG):
C
      MERR=1
      GOTO 50
  40  CONTINUE
      MERR=-1
  50  CONTINUE
C
C        >> RESET ORIGINAL "MS" VALUES:
C
      NC=2
      NMAX=NI
      MS(1,1,1)=17
      MS(1,1,2)=1
      DO 60 J=1,10
      DO 60 K=1,2
         MS(NC,J,K)=MI(J,K)
  60  CONTINUE
      RETURN
      END
C
C  *********************************************************************
C  *                                                                   *
C  *          LOADS UNIFAC BINARY INTERACTION PARAMETERS               *
C  *                                                                   *
C  *********************************************************************
      SUBROUTINE BINPAR(MDL,MGSG,AI,RI,QI,FMW,FVB)
C  *********************************************************************
      IMPLICIT  REAL*8(A-H,O-Z)
      PARAMETER  (LA=32, MA=58, NA=116)
      DIMENSION  AI(MA,MA),RI(NA),QI(NA),MGSG(NA),IJTR(NA),
     &           FMW(NA),FML(NA),FVB(NA),BB(LA)
C
      OPEN(9,FILE='RANDQ.DAT',STATUS='UNKNOWN',
     &       ACCESS='SEQUENTIAL',FORM='FORMATTED')
      DO 10 J=1,NA
         READ(9,*) MGSG(J),IJTR(J),RI(J),QI(J),FMW(J),FML(J),FVB(J)
  10  CONTINUE
C
      IF(MDL-2) 20,30,70
  20  CONTINUE
      OPEN(10,FILE='AVLE.DAT',STATUS='UNKNOWN',
     &        ACCESS='SEQUENTIAL',FORM='FORMATTED')
      GOTO 80
C
  30  CONTINUE
      OPEN(10,FILE='ALLE.DAT',STATUS='UNKNOWN',
     &        ACCESS='SEQUENTIAL',FORM='FORMATTED')
         DO 40 J=1,MA
         DO 40 K=1,MA
            AI(J,K) = 99999.D0
            IF(K.NE.J) GOTO 40
            AI(J,J) = 0.D0
  40     CONTINUE
         DO 50 J=1,LA
            READ(10,*) (BB(L),L=1,LA)
         DO 50 K=1,LA
            IF(K.EQ.J) GOTO 50
            AI(IJTR(J),IJTR(K))=BB(K)
  50     CONTINUE
CC       DO 60 K=1,NA
CC          FMW(K)=FML(K)
CC60     CONTINUE
      GOTO 100
C
  70  CONTINUE
      OPEN(10,FILE='AENV.DAT',STATUS='UNKNOWN',
     &        ACCESS='SEQUENTIAL',FORM='FORMATTED')
  80  CONTINUE
      DO 90 I=1,MA
         READ(10,*) (AI(I,J),J=1,MA)
  90  CONTINUE
 100  CONTINUE
      CLOSE(9)
      CLOSE(10)
      RETURN
      END
