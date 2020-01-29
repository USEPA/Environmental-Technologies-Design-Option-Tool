CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C																		C
C					 ADOX Model -- adoxtis								C
C																		C
C Programer:	Shumin Hu													C
C Date:		July 1998 													C
C																		C
C Description:This model can simulate a completely mixed batch reactor  	C
C			(CMBR) or a flow reactor, of which the hydraudynamics		C
C			are modeled with Tanks-in-series (TIS) completely mixed		C
C			flow reactoar (CMFR)										C
C																		C
C			The main program of this model is ADOXTIS.F.				C
C			The data are input from MODEL.DAT, and the output data		C
C			are writen in MODEL.OUT.									C
C			PHOTORATE.F and ODEQUATN.F are called in DIFFUN.F to		C
C			calculate photolysis rates and to write ordinary			C
C			differential equations.										C
C			The minimization algorithm DBCLSF from IMSL libary is used	C
C			to solve the charge balance equation to get new pH after	C
C			each simulating time step, and FCN.F is used in DBCLSF		C
C																		C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C																		C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C     NOTATION:															C
C     YO(1) = [H2O2]														C
C     YO(2) = [CHEMICAL 1 -- R1]											C
C     YO(3) = [CHEMICAL 2 -- R2]											C
C     YO(4) = [HO2*]														C
C     YO(5) = [HCO3-]														C
C     YO(6) = [H2PO4-]													C
C     YO(7) = NATURAL ORGAINC MATTER -- NOM]								C
C     YO(8) = [CO3*-]														C
C     YO(9) = [HPO4*-]													C
C     YO(10) = [OH*]														C
C																		C
C     EPS ---- DGEAR ERROR CRITEERIA										C
C     IDREACT ---- REACTOR INDENTIFICATION NUMBER							C
C		   0 -- CMBR, 1 -- CMFR											C
C     NTANK --- TANK NUMBER												C
C     VOLUME ---- TANK VOLUME, L											C
C     TAU ---- HYDRAULIC RETENSION TIME OF A TANK, min					C
C     XNTIME ---- TIMES OF TOTAL TAU THAT WILL BE SIMULATED FOR CMFR		C
C     TTOTAL ---- TOTAL SIMULATION TIME, min								C
C     SSIZE ---- SIMULATION TIME STEP, sec								C
C     OPSIZE ---- TIME INTERVAL FOR OUTPUT DATA, min						C
C     IDCARBN ---- HOW CARBONATE CONCENTRATION IS INPUT					C
C                  1 -- ALKLINITY AS CaCO3, mg/L							C
C                  0 -- TOTAL CARBONATE CONCENTRATION, M					C
C     ALK ---- ALKALINITY OF THE WATER (WHEN IDCARBN = 1)					C
C     TICARBN ---- TOTAL CARBONATE ION CONCENTRATION, M					C
C                  (WHEN IDCARBN = 0)										C
C     PH(0) ---- INITIAL pH VALUE (WHEN IDPH = 1)							C
C     PHOSPH ---- TOTAL PHOSPHATE ION CONCENTRATION, M					C
C																		C
C     NWVLEN ---- NUMBER OF WAVELENGTHS FOR LIGHT INTENSITY MEASUREMENT	C
C		        19 in the sample simulation								C
C     NIRREV ---- NUMBER OF IRRERSIBLE REACTIONS							C
C                 32 in the sample simulation								C
C     NMONOACID ---- NUMBER OF MONOPROTIC ACIDS							C
C                    equals 1+NTARGET+1; 4 in the sample simulation		C
C     NMULTIADID ---- NUMBER OF MULTIPROTIC ACIDS							C
C                    2 in the sample simulation							C
C     NREV ---- NUMBER OF REVERSIBLE REACTIONS							C
C                 8 in the sample simulation								C
C                 1:  H2O2 <==> (H+) + (HO2-)								C
C                 2:  R1H <==> (H+) + (R1-)								C
C                 3:  R2H <==> (H+) + (R2-)								C
C                 4:  HO2* <==> (H+) + (O2*-)								C
C                 5:  (HCO3-) <==> (H+) + (CO3--)							C
C                 6:  (H2PO4-) <==> (H+) + (HPO4--)						C
C                 7:  H2CO3 <==> (H+) + (HCO3-)							C
C                 8:  H3PO4 <==> (H+) + (H2PO4-)							C
C     NPHOT ---- NUMBER OF PHOTOLYSIS REACTIONS							C
C                 3, or 4 in the sample simulation						C
C     NCOMP ---- NUMBER OF COMPOUNDS INVOLVED IN THE REACTIONS			C
C                usualy it equals to 2*NMONOACID+3*NMULTIACID+4			C
C			   when CO3, PO4, and NOM present							C
C                1 -- H2O2,	2 -- R1,	3 -- R2							C
C                4 -- HO2*,	5 -- HCO3-,	6 -- H2PO4-						C
C                7 -- NOM,	8 -- CO3*-,	9 -- HPO4*-						C
C                10 -- HO*,	11 -- HO2-,	12 -- R1-						C
C                13 -- R2-,	14 -- O2*-,	15 -- CO3--						C
C                16 -- HPO4--,	17 -- H2CO3,	18 -- H3PO4				C
C	NTARGET ---- NUMBER OF TARGET ORGANIC COMPOUNDS						C
C				 2 in the sample simulation								C
C     NODE ---- NUMBER OF ORDINARY DIFFERENTIAL EQUATIONS					C
C			  equals NMONOACID+NMULTIACID+4								C
C                10 in the sample simulation								C
C                1 -- H2O2,	2 -- R1,	3 -- R2							C
C                4 -- HO2*,	5 -- HCO3-,	6 -- H2PO4-						C
C                7 -- NOM,	8 -- CO3*-,	9 -- HPO4*-						C
C                10 -- HO*												C
C     N ---- TOTAL NUMBER OF ODE											C
C																		C
C     COMNAME ---- NAME OF COMPOUNDS (20 CHARACTERS AT MOST)				C
C     CONCINI(NCOMP) ---- INITIAL CONCENTRATIONS OF THE COMPOUNDS			C
C     VALENCE(NCOMP) ---- CHARGE VALENCE OF THE COMPOUNDS					C
C     MW(NCOMP) ---- MOLECULAR WEIGHT OF THE COMPOUNDS					C
C																		C
C     XK(NIRREV) ---- REACTION RATE CONSTANTS OF IRRERSIBLE REACTIONS		C
C     XKE(NREV) ---- EQUILIBRIUM CONSTANTS OF REVERSIBLE REACTIONS		C
C     																	C
C     IRREVERSIBLE REACTION: A + B ----> C + D							C
C     COMPA(NIRREV) ---- COMPONENT INDECES OF REACTANT A'S				C
C     COMPB(NIRREV) ---- COMPONENT INDECES OF REACTANT B'S				C
C     COMPC(NIRREV) ---- COMPONENT INDECES OF REACTANT C'S				C
C     COMPD(NIRREV) ---- COMPONENT INDECES OF REACTANT D'S				C
C																		C
C     REVERSIBLE REACTION: E <====> (H+) + F								C
C     COMPE(NREV) ---- COMPONENT INDECES OF REACTANT E'S					C
C     COMPF(NREV) ---- COMPONENT INDECES OF REACTANT F'S					C
C																		C
C     PHOTOLYSIS REACTION: G ----> h H									C
C     COMPG(NPHOT) ---- COMPONENT INDECES OF REACTANT G'S					C
C     COMPF(NPHOT) ---- COMPONENT INDECES OF REACTANT F'S					C
C     STOCPHOT(NPHOT) ---- STOICHMETRIC NUMBER h FOR PHOTOLYSIS			C
C																		C
C     LWAVE(NWVLEN) ---- WAVELENGTH RANGE, nm								C
C     UVI(NWVLEN) ---- LIGHT INTENSITY AT DIFFERENT WAVELENGTH,eins./L-s	C
C     EXTCOEF(NPHOT,NWVLEN) ---- EXTINCTION COEFFICIENT OF THE COMPOUNDS	C
C				 AT DIFFERENT WAVELENGTH, 1/M-cm						C
C     QUATYD(NPHOT,NWVLEN) ---- QUATIUM YIELD OF THE COMPOUNDS			C
C	                        AT DIFFERENT WAVELENGTH						C
C     UVPATHL ---- OPTICAL PATH LENGTH OF UV-LIGHT, cm					C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C
      IMPLICIT NONE
C
      INTEGER NFLAG
      INTEGER*4 N_PW,ERROR
      PARAMETER (N_PW=750**2+2*750)
C----- Define Workspace----------------------------
C      the HUGE argument should not be needed since the DLL is
C      compiled using the /AH option.
C---For Microsoft FORTRAN
C      DOUBLE PRECISION PW[ALLOCATABLE,HUGE](:)
C---For Lahey FORTRAN
      DOUBLE PRECISION PW(:)
      ALLOCATABLE :: PW
C---- End Define Workspace-------------------------

      INTEGER MAXNTANK,MAXNTARGET,MAXEQUATN
      INTEGER MAXIRREV,MAXPHOT,MAXREV,MAXODE,MAXWVLEN,MAXCOMP,MAXSTEPS
C
	PARAMETER (MAXNTARGET=10) 
      PARAMETER (MAXNTANK=25)
      PARAMETER (MAXIRREV=100)
      PARAMETER (MAXPHOT=20)
      PARAMETER (MAXREV=20)
      PARAMETER (MAXODE=30)
      PARAMETER (MAXWVLEN=100)
      PARAMETER (MAXCOMP=50)
      PARAMETER (MAXSTEPS=2000)
	PARAMETER (MAXEQUATN=MAXODE*MAXNTANK)
C
	INTEGER I, J, K, NT, NITANK 
      INTEGER N,MF,INDEX,IDREACT,IDCARBN,IDUVI,NSTEPS,NWVLEN
      INTEGER NIRREV,NREV,NPHOT,NCOMP,NODE,NTANK
	INTEGER NTARGET,NMONOACID,NMULTIACID
	INTEGER NCARBN(MAXNTARGET),NSUBSTT(MAXNTARGET)
	INTEGER LAST_ROW
C
      DOUBLE PRECISION YO(MAXEQUATN)
      DOUBLE PRECISION XK(MAXIRREV),XKE(MAXREV)
      DOUBLE PRECISION LWAVE(MAXWVLEN),
     +          EXTCOEF(MAXPHOT,MAXWVLEN),QUATYD(MAXPHOT,MAXWVLEN)
      DOUBLE PRECISION ELECTR_POWER,UVI(MAXWVLEN),UVWATTS(MAXWVLEN),
	+                 UVEFF(MAXWVLEN)    
      DOUBLE PRECISION CONCINI(MAXCOMP)
      DOUBLE PRECISION VALENCE(MAXCOMP),MW(MAXCOMP)
C      
      INTEGER COMPA(MAXIRREV),COMPB(MAXIRREV),
     +          COMPC(MAXIRREV),COMPD(MAXIRREV),
     +          COMPE(MAXREV),COMPF(MAXREV),
     +          COMPG(MAXPHOT),COMPH(MAXPHOT)
c
      DOUBLE PRECISION STOCPHOT(MAXPHOT)
      DOUBLE PRECISION TIME(0:MAXSTEPS),TIMEMIN(0:MAXSTEPS),
     +                 CONC(MAXEQUATN+MAXNTANK,0:MAXSTEPS)
C
	DOUBLE PRECISION UNKNOWN_ION
      DOUBLE PRECISION TAU,VOLUME,UVPATHL
      DOUBLE PRECISION EPS,SSIZE,TTOTAL,XNTIMES,OPSIZE
      DOUBLE PRECISION ALK,TICARBN,PHOSPH
	DOUBLE PRECISION PH(MAXNTANK,0:MAXSTEPS),HYG(MAXNTANK)
      DOUBLE PRECISION TO,HO,TOUT
C
      CHARACTER *12 COMNAME(MAXCOMP)
C
      COMMON /DATA1/ IDREACT,TAU,NTANK
	COMMON /DATA2/ NMONOACID,NMULTIACID,XKE,UNKNOWN_ION
      COMMON /DATA3/ CONCINI,CONC,HYG,XK,NODE,NIRREV,
     +               COMPA,COMPB,COMPC,COMPD,COMPE,COMPF
C
	COMMON /PHOTO1/ NWVLEN,NPHOT,UVPATHL,EXTCOEF,QUATYD,UVI,STOCPHOT
	COMMON /PHOTO2/ COMPG,COMPH
C
	COMMON /CHARGE/NTARGET,NCARBN,NSUBSTT,PH

C
      INTEGER MAXPLOT,NUMPLOTS
      PARAMETER (MAXPLOT=100)
      CHARACTER*80 FN_INPUT,FN_OUTPUT,FN_PLOT(MAXPLOT)
C
      NFLAG=0
C
C-----Workspace allocation 
C-----For Microsoft FORTRAN
      ALLOCATE (PW(N_PW),STAT=error)
C-----For Lahey FORTRAN
C      ALLOCATE (STAT=error,PW(N_PW))
      IF (ERROR.NE.0) GOTO 9999
C-----End Workspace allocation
C
      OPEN(4,FILE='adoxpath.txt',STATUS='UNKNOWN')
      READ(4,*) FN_INPUT
      READ(4,*) FN_OUTPUT
      READ(4,*) NUMPLOTS
      DO I=1, NUMPLOTS
        READ(4,*) FN_PLOT(I)
      ENDDO
      CLOSE (4) 

      OPEN(4,FILE=fn_input,STATUS='UNKNOWN')
C
      write(*,*) "read input data from file model.dat"
	READ(4,*) NTARGET
      READ(4,*) EPS
      READ(4,*) IDREACT
      READ(4,*) NTANK
      READ(4,*) VOLUME
      READ(4,*) SSIZE
      READ(4,*) TTOTAL
      READ(4,*) XNTIMES
      READ(4,*) TAU
      READ(4,*) OPSIZE
      READ(4,*) IDCARBN
      READ(4,*) ALK
      READ(4,*) TICARBN
      READ(4,*) PH(1,0)
      READ(4,*) PHOSPH
C
      READ(4,*) UVPATHL 
      READ(4,*) NWVLEN
      DO I = 1, NWVLEN
          READ(4,*) LWAVE(I)
      ENDDO
C
	READ(4,*) ELECTR_POWER
	READ(4,*) IDUVI
	IF (IDUVI .EQ. 0) THEN
        DO I = 1, NWVLEN
          READ(4,*) UVI(I)
        ENDDO
	ELSE IF (IDUVI .EQ. 1) THEN
	  DO I = 1, NWVLEN
          READ(4,*) UVWATTS(I)
		UVI(I)=(UVWATTS(I)*LWAVE(I)*1.0D-9/0.120D0)/VOLUME
        ENDDO
	ELSE IF (IDUVI .EQ. 2) THEN
	  DO I = 1, NWVLEN
          READ(4,*) UVEFF(I)
		UVI(I)=(ELECTR_POWER*UVEFF(I)*LWAVE(I)*1.0D-9/0.120D0)/VOLUME
	  ENDDO
	ENDIF
C
      READ(4,*) NCOMP
      DO I = 1, NCOMP
          READ(4,*) COMNAME(I)
          READ(4,*) CONCINI(I)
          READ(4,*) VALENCE(I)
          READ(4,*) MW(I)
      ENDDO    
C
      READ(4,*) NIRREV
      DO I = 1, NIRREV 
          READ(4,*) COMPA(I)
          READ(4,*) COMPB(I)
          READ(4,*) COMPC(I)
          READ(4,*) COMPD(I)
          READ(4,*) XK(I)
      ENDDO
C
	NMONOACID = NTARGET + 2
      READ(4,*) NMULTIACID
      NREV = NMONOACID + 2 * NMULTIACID
      DO I = 1, NREV
          READ(4,*) COMPE(I)
          READ(4,*) COMPF(I)
          READ(4,*) XKE(I)
      ENDDO
C
      READ(4,*) NPHOT
      DO I = 1, NPHOT
          READ(4,*) COMPG(I)
          READ(4,*) COMPH(I)
          READ(4,*) STOCPHOT(I)
          DO J = 1, NWVLEN
              READ(4,*) EXTCOEF(I,J)
          ENDDO
          DO J = 1, NWVLEN
              READ(4,*) QUATYD(I,J)
          ENDDO
      ENDDO
C
	DO I = 1, NTARGET
		READ(4,*)NCARBN(I)
		READ(4,*)NSUBSTT(I)  
      ENDDO
C
C-----number of ODEs:	
C-----monoprotic acids, multiprotic acids, NOM, CO3*-, HPO4*-, HO*
C
	NODE = NMONOACID + NMULTIACID + 4
C
      IF (NTANK.GT.MAXNTANK) THEN
        PRINT *, 'Tank number exceeds maximum of ', MAXNTANK
        STOP
      ENDIF
      IF (NODE.GT.MAXODE) THEN
        PRINT *, 'Number of ODEs exceeds maximum of ', MAXODE
        STOP 
      ENDIF
      IF (NCOMP.GT.MAXCOMP) THEN
        PRINT *, 'Number of components exceeds maximum of ', MAXCOMP
        STOP 
      ENDIF
      IF (NIRREV.GT.MAXIRREV) THEN
        PRINT *, 'Number of irreversible reactions exceeds maximum of ', 
     &           MAXIRREV
        STOP 
      ENDIF
      IF (NREV.GT.MAXREV) THEN
        PRINT *, 'Number of reversible reactions exceeds maximum of ', 
     &           MAXREV
        STOP 
      ENDIF
      IF (NWVLEN.GT.MAXWVLEN) THEN
        PRINT *, 'Number of wavelengths exceeds maximum of ', MAXWVLEN
        STOP 
      ENDIF
      IF (NPHOT.GT.MAXPHOT) THEN
        PRINT *, 'Number of photolysis reactions exceeds maximum of ', 
     &           MAXPHOT
        STOP 
      ENDIF
C
      CLOSE(4)
C
      write(*,*) "end of reading"
C
      DO I = 1, NREV
        XKE(I)=10**(-XKE(I))
      ENDDO
C
	DO I = 1, NTANK
        HYG(I)=10**(-PH(1,0))
	ENDDO
C
      IF (IDCARBN.EQ.1) THEN
        TICARBN = ALK/50.0
      ENDIF
C
      CONCINI(NMONOACID+1) = TICARBN
      CONCINI(NMONOACID+2) = PHOSPH
C
C-----calculate the total number of ODE
C
	IF (IDREACT.EQ.0) THEN
	  NTANK = 1
	ENDIF
      N = NTANK * NODE
C
C-----assignment of initial concentration, they are "total conc."
C
      TO=0.0
	DO I = 1, NTANK
        DO J = 1, NODE
          YO((I-1)*NODE+J)= CONCINI(J)
	  ENDDO
      ENDDO
C-----convert "total conc," into that of disassociated form
	DO I = 1, NTANK
        DO J = 1, NMONOACID
          YO((I-1)*NODE+J)=YO((I-1)*NODE+J)/(1.0+XKE(J)/HYG(1))
        ENDDO
        DO J = 1, NMULTIACID
          YO((I-1)*NODE+NMONOACID+J)=YO((I-1)*NODE+NMONOACID+J)/(1.0
     +          +HYG(1)/XKE(NMONOACID+NMULTIACID+J)
     +          +XKE(NMONOACID+J)/HYG(1))
	  ENDDO
      ENDDO
C
C-----end of initial concentration assignment
C
C-----calculate the unknown charged species from initial charge balance
C-----UNKNOWN_ION=(2*[CO3--]+[HCO3-]+[OH-]+[HO2-]+[R1-]+[R2-]+...)-[H+]
C
	UNKNOWN_ION=10**(-14+PH(1,0))-HYG(1)
	DO I = 1, NMONOACID
	    UNKNOWN_ION=UNKNOWN_ION+YO(I)*XKE(I)/HYG(1)
	ENDDO
	DO I = NMONOACID+1, NMONOACID + NMULTIACID
          UNKNOWN_ION=UNKNOWN_ION+YO(I)+2*YO(I)*XKE(I)/HYG(1)
	ENDDO
C
C-----end of UNKNOWN_ION calculation
C
      IF (IDREACT.EQ.1) THEN
	  TAU=60.0*TAU
        TTOTAL=XNTIMES*TAU
	ELSE
	  TTOTAL=60.0*TTOTAL
      ENDIF
C
      NSTEPS=TTOTAL/SSIZE+1
C
      IF (NSTEPS.GT.MAXSTEPS) THEN
        PRINT *, 'Number of time steps exceeds maximum of ', MAXSTEPS
        STOP 
      ENDIF
C
      HO=1.0D-9
      MF=22
      INDEX=1
C
C-----assign the initial concentration to CONC(MAXEQUATN, NSTEPS)
C-----for the purpose of output, they are "total conc."
C
      TIME(0)=0
      TIMEMIN(0)=0
C
	DO I = 1, NODE
	  CONC(I,0)=CONCINI(I)
	ENDDO
	DO I = 2, NTANK
	  DO J = 1, NODE
          CONC((I-1)*NODE+J,0)=0
	  ENDDO
      ENDDO
C
	DO NT = 1, NTANK
	  CONC(N+NT,0)=PH(1,0)
	ENDDO
C
C-----store initial yo(I) value of first tank back in "initial conc." for modeling
C-----flow reactor, they are not "total conc.", but that of the disassociated form
C 
      DO I = 1, NODE
          CONCINI(I) = YO(I)
      ENDDO
C
C-----call Dgear subroutine
C
      TOUT= 0.0
C
      DO 1000 K = 1,NSTEPS
C
        TOUT=TOUT+SSIZE
        TIME(K)=TOUT
        TIMEMIN(K)=TOUT/60.0
C
        CALL DGEAR (N,TO,HO,YO,TOUT,EPS,MF,INDEX,NFLAG,PW,N_PW)
C
	  WRITE(*,*) '% complete ', (100.0D0*TOUT)/TTOTAL
C
	  DO I = 1, NTANK
	    PH(I,K)=-LOG10(HYG(I))
	  ENDDO
C
C-----assign the calculated concentration to CONC(MAXEQUATN, NSTEPS) for output
C-----the output concentrations are the total value of both associated and
C-----disassociated forms, e.g. [H2O2]t = [H2O2] + [HO2-]
C
        DO I = 1, N
          CONC(I,K)=YO(I)
        ENDDO
C-----convert the disassociated form of conc. into to "total conc."
	  DO J = 1, NTANK
          DO I = (J-1)*NODE+1, (J-1)*NODE+NMONOACID
              CONC(I,K)=CONC(I,K)*(1.0+XKE(I-(J-1)*NODE)/HYG(J))
          ENDDO
          DO I = (J-1)*NODE+NMONOACID+1, (J-1)*NODE+NMONOACID+NMULTIACID
              CONC(I,K)=CONC(I,K)*(1.0
     +          +HYG(J)/XKE(NMULTIACID+I-(J-1)*NODE)
     +          +XKE(I-(J-1)*NODE)/HYG(J))
          ENDDO
		CONC(N+J,K)=PH(J,K)
	  ENDDO
C
1000  CONTINUE
C
C-----write output files: model.out and plot files
C
	DO I = 1, NODE
	  WRITE(*,*) '- Writing to ', fn_plot(i)
	  OPEN(8,FILE=fn_plot(i),STATUS='UNKNOWN') 
        DO NT = 1, NTANK
	    NITANK = (NT-1)*NODE
	    DO J = 0, NSTEPS,OPSIZE*60/SSIZE
	      WRITE(8,100) NT, TIME(J),TIMEMIN(J),CONC(NITANK+I,J)
	    ENDDO
	  ENDDO
	  CLOSE (8)
	ENDDO
C
      OPEN(7,FILE=fn_output,STATUS='UNKNOWN')
      WRITE(*,*) '- Writing to ', fn_output
C
	DO NT = 1, NTANK
	  NITANK = (NT-1)*NODE
        DO I = 1, (NODE-1)/5
          WRITE(7,50) COMNAME(5*I-4),COMNAME(5*I-3),
	+                COMNAME(5*I-2),
     +                COMNAME(5*I-1),COMNAME(5*I)
          DO J = 0, NSTEPS,OPSIZE*60/SSIZE
              WRITE(7,100) NT,TIME(J),TIMEMIN(J),
	+                CONC((NITANK+(5*I-4)),J),CONC((NITANK+(5*I-3)),J),
     +                CONC((NITANK+(5*I-2)),J),CONC((NITANK+(5*I-1)),J),
     +                CONC((NITANK+(5*I)),J)
          ENDDO
        ENDDO
	  LAST_ROW = MOD(NODE,5)
	  IF (LAST_ROW .EQ. 0) THEN
		LAST_ROW = 5
	  ENDIF
	  I = (NODE-1)/5 + 1
	  IF (LAST_ROW .EQ. 1) THEN
	    WRITE(7,50) COMNAME(NODE)
		DO J = 0, NSTEPS,OPSIZE*60/SSIZE
		  WRITE(7,100) NT,TIME(J),TIMEMIN(J),CONC((NITANK+NODE),J)
		ENDDO
	  ENDIF
	  IF (LAST_ROW .EQ. 2) THEN
	    WRITE(7,50) COMNAME(NODE-1),COMNAME(NODE)
		DO J = 0, NSTEPS,OPSIZE*60/SSIZE
		  WRITE(7,100) NT,TIME(J),TIMEMIN(J),CONC((NITANK+NODE-1),J),
	+              CONC((NITANK+NODE),J)
		ENDDO
	  ENDIF
	  IF (LAST_ROW .EQ. 3) THEN
	    WRITE(7,50) COMNAME(NODE-2),COMNAME(NODE-1),COMNAME(NODE)
		DO J = 0, NSTEPS,OPSIZE*60/SSIZE
		  WRITE(7,100) NT,TIME(J),TIMEMIN(J),CONC((NITANK+NODE-2),J),
	+              CONC((NITANK+NODE-1),J),CONC((NITANK+NODE),J)
		ENDDO
	  ENDIF
	  IF (LAST_ROW .EQ. 4) THEN
	    WRITE(7,50) COMNAME(NODE-3),COMNAME(NODE-2),COMNAME(NODE-1),
	+                COMNAME(NODE)
		DO J = 0, NSTEPS,OPSIZE*60/SSIZE
		  WRITE(7,100) NT,TIME(J),TIMEMIN(J),CONC((NITANK+NODE-3),J),
	+              CONC((NITANK+NODE-2),J),CONC((NITANK+NODE-1),J),
	+              CONC((NITANK+NODE),J)
		ENDDO
	  ENDIF
	  IF (LAST_ROW .EQ. 5) THEN
	    WRITE(7,50) COMNAME(NODE-4),COMNAME(NODE-3),COMNAME(NODE-2),
	+                COMNAME(NODE-1),COMNAME(NODE)
		DO J = 0, NSTEPS,OPSIZE*60/SSIZE
		  WRITE(7,100) NT,TIME(J),TIMEMIN(J),CONC((NITANK+NODE-4),J),
	+              CONC((NITANK+NODE-3),J),CONC((NITANK+NODE-2),J),
	+              CONC((NITANK+NODE-1),J),CONC((NITANK+NODE),J)
		ENDDO
	  ENDIF
	ENDDO
C
	WRITE(7,150)
	DO NT = 1, NTANK
        DO J = 0, NSTEPS,OPSIZE*60/SSIZE
	    WRITE(7,200) NT,TIME(J),TIMEMIN(J),CONC(N+NT,J)
	  ENDDO
	ENDDO
      CLOSE (7)
C
50    FORMAT('#TANK',T7,'TIME(SEC)',T17,'TIME(MIN)',T27,
	+      A12,A12,A12,A12,A12)
100   FORMAT(1X,I2,T5,F7.0,T15,F5.1,T24,E9.3E2,T36,E9.3E2,T48,E9.3E2,
     +       T60,E9.3E2,T72,E9.3E2)
150   FORMAT('#TANK',T7,'TIME(SEC)',T17,'TIME(MIN)',T27,'pH')
200   FORMAT(1X,I2,T5,F7.0,T15,F5.1,T24,f5.2)
      WRITE(*,*) "finish writing"
C     
9999  IF (ERROR.NE.0) then 
        NFLAG = 1603 
      ELSE
C-----For Microsoft FORTRAN
       DEALLOCATE(PW,STAT=error)
C-----For Lahey FORTRAN
C      DEALLOCATE(STAT=ERROR,PW)
        IF (ERROR.NE.0) THEN 
          NFLAG = 1603
        ENDIF
      ENDIF
C
C     END OF THE MAIN PROGRAM
C
      STOP
      END