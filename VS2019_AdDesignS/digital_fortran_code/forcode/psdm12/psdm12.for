CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
CC
CC  Program Name:       PSDM for DLL
CC  Author:             Michigan Tech University - 1994
CC  Intended Platform:  Compiled with Microsoft FORTRAN and linked
CC                      to the Visual Basic code of the Adsorption
CC                      Design Software (AdDesignS).
CC
CC  Modification History:
CC  =====================
CC  11/18/1994: Fred Gobin
CC  - User input for the diffusion coefficients
CC  - Valid gas and liquid phase
CC  - 6 component maximum
CC  - Uses DGEAR subroutine
CC  - Fouling correlation for K reduction
CC  - Time variable tortuosity included (see DIFFUN)
CC  03/16/1996: Eric Oman
CC  - User can now input tortuosity for every component
CC  04/18/1996: Eric Oman
CC  - Modified to send some input/calculated variables back to caller
CC    - VARS1: General simulation variables
CC    - VARS2: Component-specific simulation variables
CC  07/04/1996: Eric Oman
CC  - Modified to output PSDM internal variables to file called PSDM.XO,
CC    if parameter DEBUGM is set to 1.
CC  07/19/1996: Eric Oman
CC  - Modified to use Bhu's code for variable input, initialization,
CC    DIFFUN, ORTHOG, etc.
CC  02/08/1997: Dave Hokanson
CC  - Modified to have the capability to use up to 18 radial 
CC    collocation constants.  At this point, maximum number of
CC    equations to by solved at DGEAR will still be left at 750
CC    in the Visual Basic code.
CC    Later, Y0, YDOT and workspace sizes may be increased to use
CC    18 axial collocation points and 18 radial collocation points
CC    at the same time.
CC  12/09/1997: Eric Oman
CC  - Removed debug output files.
CC
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
CC
CC  Description of Inputs/Outputs:
CC  ==============================
CC
CC  *I*   Ads_Prop    Adsorber Bed Properties
CC                    Array Size: (4)
CC                    (1): Length (m)
CC                    (2): Diameter (m)
CC                    (3): Weight of adsorbent (kg)
CC                    (4): Inlet flowrate (m^3/s)
CC  *I*   C_Prop      Adsorbent Properties
CC                    Array Size: (3)
CC                    (1): Void Fraction of the particle (-)
CC                    (2): Apparent Density (g/cm^3)
CC                    (3): Particle Radius (cm)
CC  *I*   Chemicals   Chemical Properties
CC                    Array Size: (Numb,16)
CC                    (I,1): MW (g/mol)
CC                    (I,2): Initial conc. (ug/l)
CC                    (I,3): Molar Volume (cm^3/mol)
CC                    (I,4): Freundlich K (*)
CC                    (I,5): Freundlich 1/n
CC                    (I,6): kf (cm/s)
CC                    (I,7): Ds (cm^2/s)
CC                    (I,8): Dp (cm^2/s)
CC                    (I,9): Coeff. for fouling correlation (-)
CC                    (I,10): Coeff. for fouling correlation (1/min)
CC                    (I,11): Coeff. for fouling correlation (-)
CC                    (I,12): Coeff. for fouling correlation (1/min)
CC                    (I,13): Tortuosity
CC                              Note: This input is never used!
CC                              - ejo, 3/16/96
CC                    (I,14): Tor. Coeff. for Tortuosity=f(t)
CC                    (I,15): Part. Coeff. for Tortuosity=f(t)
CC                    (I,16): Time parameter(min) for Tortuosity=f(t)
CC  *I*   CinI        Influent Concentrations (ug/L)
CC                    Array Size: (Numb,NinI)
CC  *O*   CPVB        Reduced Breakthrough Concentrations, C/C0 (-)
CC                    Array Size: (Numb,400)
CC  *I*   ISDBUG      Debug mode for program
CC                    Setting of 0 ===> No debugging
CC                    Setting of 1 ===> Outputs to various text files
CC  *I*   MXX         Number of axial collocation points
CC  *I*   N_PW        Size of the working space (bytes)
CC  *O*   NFLAG       Error flag returned to Visual Basic
CC  *I*   NinI        Number of influent points (see CinI and TinI)
CC  *O*   NITP        Number of breakthrough points (see CPVB and T)
CC  *I*   Numb        Number of chemicals
CC  *I*   NumBed      Current bed number in series to simulate
CC  *I*   NXX         Number of radial collocation points
CC  *O*   T           Breakthrough Times (minutes)
CC                    Array Size: (400,2)
CC  *I*   TinI        Times for CinI array (minutes)
CC                    Array Size: (NinI)
CC  *I*   TT          Time parameters
CC                    Array Size: (3)
CC                    (1): Time to end simulation (minutes)
CC                    (2): Time to begin simulation (minutes)
CC                    (3): Time step (minutes)
CC  *O*   VARS1       Various debugging variables
CC                    Array Size: (15)
CC  *O*   VARS2       Various debugging variables
CC                    Array Size: (Numb,19)
CC
CC (*) --- K is in (umol/g)x(L/umol)^(1/n)
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC



C      SUBROUTINE PSDM(Numb,Chemicals,Ads_Prop,C_Prop,
C     &                T,CPVB,NITP,TT,NXX,MXX,
C     &                NinI,TinI,CinI,N_PW,NumBed,NFLAG,
C     &                VARS1,VARS2,ISDBUG)
      SUBROUTINE PSDM(Numb,Chemicals,Ads_Prop,C_Prop,TT,
     &    MXX,NXX,N_PW,NinI,TinI,CinI,NumBed,
     &    VARS1,VARS2,NITP,T,CPVB,NFLAG)
    
      IMPLICIT NONE
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
C
C**** Define Workspace **********************************************
c     The HUGE argument should not be needed since the DLL is
c     compiled using the /AH option.
      DOUBLE PRECISION PW[ALLOCATABLE,HUGE](:)

C
C------ PARAMETERS TO CALCULATION MODULE.
C
C...INPUTS...:
      INTEGER*2 Numb
      DOUBLE PRECISION Chemicals(MXCOMP,16)
      DOUBLE PRECISION Ads_Prop(4)
      DOUBLE PRECISION C_Prop(3)
      DOUBLE PRECISION TT(3)
      INTEGER*2 MXX
      INTEGER*2 NXX
      INTEGER*4 N_PW
      INTEGER*2 NinI
      DOUBLE PRECISION TinI(MAXPTS)
      DOUBLE PRECISION CinI(MXCOMP,MAXPTS)
      INTEGER*2 NumBed
C...OUTPUTS...:
      DOUBLE PRECISION VARS1(15)
      DOUBLE PRECISION VARS2(MXCOMP,19)
      INTEGER*2 NITP
      DOUBLE PRECISION T(MAXPTS,2)
      DOUBLE PRECISION CPVB(MXCOMP,MAXPTS)
      INTEGER*2 NFLAG
C      INTEGER*2 Numb
C      DOUBLE PRECISION Chemicals(Numb,16)
C      DOUBLE PRECISION Ads_Prop(4)
C      DOUBLE PRECISION C_Prop(3)
C      DOUBLE PRECISION T(MAXPTS,2)
C      DOUBLE PRECISION CPVB(Numb,MAXPTS)
C      INTEGER*2 NITP
C      DOUBLE PRECISION TT(5)
C      INTEGER*2 NXX,MXX,NinI
C      DOUBLE PRECISION TinI(NinI),CinI(Numb,NinI)
C      INTEGER*4 N_PW
C      INTEGER*2 NumBed
C      INTEGER*2 NFLAG
C      DOUBLE PRECISION VARS1(15)
C      DOUBLE PRECISION VARS2(Numb,19)
C      INTEGER ISDBUG

c---- Storage of input variables
      INTEGER*2 NCOMP
      DOUBLE PRECISION XWT(MXCOMP),CBO(MXCOMP),VB(MXCOMP),XK(MXCOMP),
     &                 XN(MXCOMP),KF(MXCOMP),DS(MXCOMP),DP(MXCOMP),
     &                 RK1(MXCOMP),RK2(MXCOMP),RK3(MXCOMP),RK4(MXCOMP),
     &                 TORTU(MXCOMP),TOR(MXCOMP),PART(MXCOMP),
     &                 TTORTU(MXCOMP)
      DOUBLE PRECISION L,DIA,WT,FLRT
      DOUBLE PRECISION EPOR,RHOP,RAD
      INTEGER*2 BEDNUM
      DOUBLE PRECISION DSTEP,DTOL,DOUT
      INTEGER*2 MC,NC
      INTEGER*2 NIN
      DOUBLE PRECISION TIN(MAXPTS),CIN(MXCOMP,MAXPTS)

c---- Storage of calculated variables
      INTEGER*2 NEQ
      DOUBLE PRECISION AREA,BEDVOL,EBED,EBCT,TAU
      INTEGER*2 NCA,MCA
      DOUBLE PRECISION AZ1(MAXMC,MAXMC),BR1(MAXNC,MAXNC),WR1(MAXNC)
      DOUBLE PRECISION AZ(MAXMC,MAXMC),BR(MAXNC,MAXNC),WR(MAXNC)
      DOUBLE PRECISION QTE
      DOUBLE PRECISION D(MXCOMP),QE(MXCOMP),DGS(MXCOMP),DGP(MXCOMP)
      DOUBLE PRECISION EDS(MXCOMP),EDP(MXCOMP),ST(MXCOMP)
      DOUBLE PRECISION BIS(MXCOMP),BIP(MXCOMP),DG(MXCOMP)
      DOUBLE PRECISION XNI(MXCOMP)
      DOUBLE PRECISION DGT,BVF
      DOUBLE PRECISION TCONV,TSTEP,TTOL
      DOUBLE PRECISION YM(MXCOMP)
      INTEGER*4 N
      DOUBLE PRECISION Y0(MAXDE)
      DOUBLE PRECISION DT0,DH0,T0,H0,TOUT,EPS
      INTEGER*2 ITP
      INTEGER*4 MF
      INTEGER*4 INDEX

c---- Major calculation loop
      DOUBLE PRECISION CP(MXCOMP,MAXPTS)
      DOUBLE PRECISION TP(MAXPTS)
      INTEGER*2 N1

c---- Miscellaneous
      INTEGER*2 I,J

c---- Debug Variables
      INTEGER*2 DEBUGM
      DOUBLE PRECISION LAST_T
      COMMON /DEBUG/ LAST_T, DEBUGM
      INTEGER*4 error

c---- Common blocks

      COMMON /BLOCKA/ DG,ST,EDS,EDP,BR,D
      COMMON /BLOCKB/ YM,XNI,XN,WR,AZ
      COMMON /BLOCKC/ MC,NC,NCOMP,N1,DGT,NIN
      COMMON /BLOCKD/ CIN,TIN
      COMMON /BLOCKF/ TOR,PART,TCONV,TORTU
      COMMON /BLOCKG/ RK1,RK2,RK3,RK4,XK
      COMMON /BLOCKH/ BEDNUM
      COMMON /BLOCKJ/ CBO

      INTEGER ISDBUG

C
C                  SET UP DEBUG VARIABLES AND OPEN DEBUG FILES
C

      ISDBUG = 0
      IF (ISDBUG .EQ. 1) THEN
	DEBUGM = 1
	LAST_T = 0D0
      ELSE
	DEBUGM = 0
      ENDIF

c      open(3,file='test.txt')
c      write(3,*) 'Got to this point (A) in PSDM10'
c      close(3)

      IF (DEBUGM .EQ. 1) THEN
	OPEN(4,FILE='psdm.xo')
	OPEN(8,FILE='psdm.y')
      END IF

C
C                  ALLOCATE WORKSPACE FOR DIFFERENTIAL EQUATIONS
C

      ALLOCATE (PW(N_PW),STAT=error)
      IF (error.NE.0) GOTO 9999

C
C                  READ IN INPUT DATA
C

c---- Input number of chemicals
      NCOMP = Numb

c---- Input chemical properties
      DO 801 I = 1,NCOMP
	XWT(I)    = Chemicals(I,1)
	CBO(I)    = Chemicals(I,2)
	VB(I)     = Chemicals(I,3)
	XK(I)     = Chemicals(I,4)
	XN(I)     = Chemicals(I,5)
	KF(I)     = Chemicals(I,6)
	DS(I)     = Chemicals(I,7)
	DP(I)     = Chemicals(I,8)
	RK1(I)    = Chemicals(I,9)
	RK2(I)    = Chemicals(I,10)
	RK3(I)    = Chemicals(I,11)
	RK4(I)    = Chemicals(I,12)
	TORTU(I)  = Chemicals(I,13)
	TOR(I)    = Chemicals(I,14)
	PART(I)   = Chemicals(I,15)
	TTORTU(I) = Chemicals(I,16)
  801 CONTINUE

c---- Input adsorption bed properties
c     Note : L,DIA are converted from meters ---> centimeters
c            WT is converted from kilograms ---> grams
c            FLRT is converted from cubic meters per second
c            ---> milliliters per minute
      L      = Ads_Prop(1)*100.D0
      DIA    = Ads_Prop(2)*100.D0
      WT     = Ads_Prop(3)*1000.D0
      FLRT   = Ads_Prop(4)*60.D0*1D6

c---- Input carbon properties
      EPOR   = C_Prop(1)
      RHOP   = C_Prop(2)
      RAD    = C_Prop(3)

c---- Input set the number of the bed in series being handled
      BEDNUM = NumBed

c---- Input set the simulation time parameters
      DSTEP  = TT(3)
      DTOL   = TT(1)
      DOUT   = TT(2)

c---- Input number of collocation points
      MC = MXX
      NC = NXX

c---- Input variable influent concentrations
      NIN = NinI
      IF (NIN .EQ. 0) GO TO 812
      DO 1 J = 1,NIN                                                    
C       READ(4,*) TIN(J), (CIN(I,J), I = 1,NCOMP)                      
	TIN(J) = TinI(J)
	DO 2 I=1,NCOMP
	  CIN(I,J) = CinI(I,J)
2       CONTINUE
1     CONTINUE
							  
C
C                  CALCULATIONS
C

c---- Calculate number of equations
812   NEQ = (MC*(NC + 1) - 1)*NCOMP

c---- Convert influent concentrations from ug/L ---> umol/L
      DO 212 I = 1, NCOMP
	CBO(I) = CBO(I)/XWT(I)
	DO 211 J=1, NIN
	  CIN(I,J) = CIN(I,J)/XWT(I)
211     CONTINUE
212   CONTINUE                                                          

c---- Calculate various bed parameters
      AREA = 3.141592654D0*DIA*DIA/4.0D0
      BEDVOL = L*AREA                                                   
      EBED = 1.0D0 - WT/(BEDVOL*RHOP)                                   
      EBCT = BEDVOL/FLRT                                                
      TAU = BEDVOL*EBED*60.0D0/FLRT                                     

c---- Calculate collocation constants
      MCA = MC
      NCA = NC
      CALL CONSTNT(NCA,MCA,AZ1,BR1,WR1,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      DO 9 I=1,MC
	DO 9 J=1,MC
	  AZ(I,J) = AZ1(I,J)
9     CONTINUE
      DO 8 I=1,NC
	WR(I)=WR1(I)
	DO 8 J=1,NC
	  BR(I,J) = BR1(I,J)
8     CONTINUE

c---- Calculate some dimensionless groups
      QTE = 0.0
      DO 30 I = 1,NCOMP                                                 
	D(I) = DS(I)/DP(I)
	QE(I) = XK(I)*CBO(I)**XN(I)
	QTE = QTE + QE(I)
	DGS(I) = (RHOP*QE(I)*(1.0 - EBED)*1000.0)/(EBED*CBO(I))
	DGP(I) = EPOR*(1.0 - EBED)/EBED
	EDS(I) = DS(I)*DGS(I)*TAU/(RAD**2)
	IF (EDS(I) .LE. 1.0D-130) THEN
	  EDS(I) = 1.0D-130
	ENDIF
	EDP(I) = DP(I)*DGP(I)*TAU/(RAD**2)
	ST(I)  = KF(I)*(1.0 - EBED)*TAU/(EBED*RAD)
	BIS(I) = ST(I)/EDS(I)
	BIP(I) = ST(I)/EDP(I)
	DG(I)  = DGS(I) + DGP(I)
	XNI(I) = 1.0/XN(I)
30    CONTINUE

c---- Calculate total solute distribution parameter, DGT,
c.... and bed volumes fed to column, BVF
      DGT = 0.0                                                         
      DO 33 I = 1,NCOMP                                                 
	 DGT = DGT + DG(I)                                              
33    CONTINUE
      BVF = EBED*DGT                                                    

c---- Calculate some time parameters
      TCONV = 60.0/(TAU*(DGT + 1))
      TSTEP = DSTEP*TCONV                                               
      TTOL  = DTOL*TCONV                                                

c---- Calculate equilibrium adsorbent phase concentration fractions
      DO 35 I = 1,NCOMP                                                 
	YM(I) = QE(I)/QTE
35    CONTINUE

c---- Call subroutine ORTHOG to combine collocation constants
c.... and dimensionless groups and to determine total number
c.... of differential equations being solved for by GEAR
      CALL ORTHOG(N)

c---- Initialize dependent variables
      DO 65 I = 1, N
	Y0(I) = 0.0
65    CONTINUE

c---- Parameters for DGEAR
      DT0   = 0.0
      DH0   = 1.0D-9

      T0    = DT0*TCONV                                                 
      H0    = DH0*TCONV                                                 
      TOUT  = DOUT*TCONV                                                
      EPS   = 1.0D-3
      MF    = 22
      INDEX = 1

c---- Convert influent and experimental data to dimensionless form
      DO 60 J = 1,NIN
	TIN(J) = TIN(J)*TCONV
	DO 55 I = 1,NCOMP
	  CIN(I,J) = CIN(I,J)/CBO(I)
55      CONTINUE
60    CONTINUE

C
C                  OUTPUT DEBUG FILES
C


      IF (DEBUGM .EQ. 1) THEN
	WRITE(4,*) '***** Before DGEAR Loop'
	WRITE(4,*) 'N =', N
	WRITE(4,*) 'T0 =', T0
	WRITE(4,*) 'H0 =', H0
	WRITE(4,*) 'Y0 =', (Y0(I), I=1, N)
	WRITE(4,*) 'TOUT =', TOUT
	WRITE(4,*) 'EPS =', EPS
	WRITE(4,*) 'MF =', MF
	WRITE(4,*) 'INDEX =', INDEX
      END IF

c      IF (DEBUGM .EQ. 1) THEN
c        OPEN(3,FILE='psdm.v1')
c        WRITE(3,*) 'QTE =', QTE
c        WRITE(3,*) 'QE(1) =', QE(1)
c        WRITE(3,*) 'YM(1) =', YM(1)
c        WRITE(3,*) '---'
c        I = 1
c        WRITE(3,*) 'I =', I
c        WRITE(3,*) 'CBO(I) =', CBO(I)
c        WRITE(3,*) 'D(I) =', D(I)
cc        WRITE(3,*) 'DEN =', DEN
cc        WRITE(3,*) 'DGI(I) =', DGI(I)
c        WRITE(3,*) 'DGP(I) =', DGP(I)
c        WRITE(3,*) 'DGS(I) =', DGS(I)
c        WRITE(3,*) 'DGT =', DGT
cc        WRITE(3,*) 'DIFL(I) =', DIFL(I)
c        WRITE(3,*) 'DP(I) =', DP(I)
c        WRITE(3,*) 'DS(I) =', DS(I)
c        WRITE(3,*) 'EBCT =', EBCT
c        WRITE(3,*) 'EBED =', EBED
cc        WRITE(3,*) 'EDD(I) =', EDD(I)
c        WRITE(3,*) 'EDP(I) =', EDP(I)
c        WRITE(3,*) 'EDS(I) =', EDS(I)
c        WRITE(3,*) 'EPOR =', EPOR
c        WRITE(3,*) 'QE(I) =', QE(I)
c        WRITE(3,*) 'RAD =', RAD
c        WRITE(3,*) 'RHOP =', RHOP
cc        WRITE(3,*) 'SPDFR =', SPDFR
cc        WRITE(3,*) 'STD(I) =', STD(I)
c        WRITE(3,*) 'TAU =', TAU
c        WRITE(3,*) 'TOR(I) =', TOR(I)
cc        WRITE(3,*) 'TORT(I) =', TORT(I)
c        WRITE(3,*) 'VB(I) =', VB(I)
cc        WRITE(3,*) 'VW =', VW
c        WRITE(3,*) 'XK(I) =', XK(I)
c        WRITE(3,*) 'XN(I) =', XN(I)
c        CLOSE(3)
c      END IF

C
C                  MAJOR CALCULATION LOOP TO INTEGRATE D.E.'s
C

      ITP = 0                                                           
   70 ITP = ITP + 1                                                     

      CALL DGEAR (N,T0,H0,Y0,TOUT,EPS,MF,INDEX,NFLAG,PW,N_PW)

      IF (ALLOW_SCREENIO.EQ.1) THEN
        WRITE (*,'(1X,''AX-ELEM '',I2,'' OF '',I2,''; ' //
     &           'PERCENT COMPLETE = '',F7.2,''%'')') 
     &           NumBed,
     &           TotalAxialElementCount_Copy,
     &           (100.0D0*TOUT)/TTOL
      ENDIF

      IF (DEBUGM .EQ. 1) THEN
	WRITE(4,*) '***** Out of DGEAR Loop'
	WRITE(4,*) 'N =', N
	WRITE(4,*) 'T0 =', T0
	WRITE(4,*) 'H0 =', H0
	WRITE(4,*) 'Y0 =', (Y0(I), I=1, N)
	WRITE(4,*) 'TOUT =', TOUT
	WRITE(4,*) 'EPS =', EPS
	WRITE(4,*) 'MF =', MF
	WRITE(4,*) 'INDEX =', INDEX
      END IF

      DO 75 I = 1,NCOMP                                                 
	CP(I,ITP) = Y0(N1*I)
75    CONTINUE
      TP(ITP) = TOUT                                                    
      DOUT = TOUT/TCONV                                                 
      IF ( ITP .LT. NSTEPS ) THEN                                       
	IF ( TOUT .LT. TTOL ) THEN
	  TOUT = TOUT + TSTEP
	  IF ( TOUT .GT. TTOL ) TOUT = TTOL
	  GOTO 70
	ENDIF
      ELSE                                                              
	IF ( TOUT .NE. TTOL ) THEN
c          WRITE(6,108) NSTEPS, DOUT
	  GOTO 81
	ENDIF
      ENDIF                                                             

C
C                  TRANSFER DATA BACK TO VISUAL BASIC
C

81    DO 82 J = 1, ITP
	T(J,1) = TP(J)/TCONV
	T(J,2) = TP(J)*BVF
	DO 821 I=1, NCOMP
	  CPVB(I,J) = CP(I,J)
821     CONTINUE
82    CONTINUE

      NITP = ITP

C
C                  DEALLOCATE D.E. WORKSPACE MEMORY
C

9999  IF (error.NE.0) then
	NFLAG = 1603
      ELSE
	DEALLOCATE(PW,STAT=error)
	IF (error.NE.0) THEN
	  NFLAG = 1603
	ENDIF
      ENDIF

C
C                  OUTPUT SOME DEBUG VARIABLES TO VISUAL BASIC
C

      VARS1(1) = NC / 1D0
      VARS1(2) = MC / 1D0
      VARS1(3) = NEQ / 1D0
      VARS1(4) = RAD
      VARS1(5) = RHOP
      VARS1(6) = EPOR
      VARS1(7) = EBED
C      VARS1(8) = ----
      VARS1(9) = TAU
      VARS1(10) = EBCT
C      VARS1(11) = ----
C      VARS1(12) = ----
C      VARS1(13) = ----
C      VARS1(14) = ----
      VARS1(15) = NFLAG / 1D0

      DO 1300 I = 1,NCOMP
	VARS2(I,1) = VB(I)
	VARS2(I,2) = XWT(I)
	VARS2(I,3) = CBO(I)
	VARS2(I,4) = XK(I)
	VARS2(I,5) = XN(I)
C        VARS2(I,6) = ----
	VARS2(I,7) = KF(I)
	VARS2(I,8) = DS(I)
	VARS2(I,9) = ST(I)
	VARS2(I,10) = DGS(I)
	VARS2(I,11) = BIS(I)
	VARS2(I,12) = EDS(I)
	VARS2(I,13) = DGP(I)
	VARS2(I,14) = DP(I)
	VARS2(I,15) = BIP(I)
	VARS2(I,16) = EDP(I)
	VARS2(I,17) = D(I)
C        VARS2(I,18) = ----
1300  CONTINUE

      IF (DEBUGM .EQ. 1) THEN
	CLOSE(4)
	CLOSE(8)
      END IF

      RETURN
      END                                                               
C                                                                       
C **********************************************************************                                                                       
		       SUBROUTINE ORTHOG ( N )                          
C **********************************************************************                                                                      
c      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

c---- Subroutine parameters
      INTEGER*4 N
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
Cc---- Constants
C      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE
C
Cc**** Change Hokanson 2/8/97
Cc      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=6,MAXPTS=400,MAXDE=750)
C      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
Cc**** End Change Hokanson 2/8/97

c---- Local variables
      DOUBLE PRECISION EDD(MXCOMP)

c---- Common block variables
      DOUBLE PRECISION DG(MXCOMP),ST(MXCOMP),EDS(MXCOMP),EDP(MXCOMP),
     &                 BR(MAXNC,MAXNC),D(MXCOMP)
      INTEGER*2 MC,NC,NCOMP,N1
      DOUBLE PRECISION DGT
      INTEGER*2 NIN
      DOUBLE PRECISION STD(MXCOMP),BEDS(MXCOMP,MAXNC,MAXNC),
     &                 BEDP(MXCOMP,MAXNC,MAXNC),
     &                 DGI(MXCOMP)
      INTEGER*2 MND,ND,MD

c---- Common block
      COMMON /BLOCKA/ DG,ST,EDS,EDP,BR,D
      COMMON /BLOCKC/ MC,NC,NCOMP,N1,DGT,NIN
      COMMON /BLOCKE/ STD,BEDS,BEDP,DGI,MND,ND,MD

c---- Debug variables
      INTEGER*2 DEBUGM
      DOUBLE PRECISION LAST_T
      COMMON /DEBUG/ LAST_T, DEBUGM

      ND  = NC - 1                                                      
      MD  = MC - 1                                                      
      MND = MC*ND                                                       
      N1  = MND + MC + MD                                               
      N   = N1*NCOMP                                                    
c---Modified by ejo on 7/19/96
c      DGT = 1.0 + DGT
c      DO 50 I = 1,NCOMP
c---Original Code Follows--Modified by ejo on 7/19/96
      DO 50 I = 1,NCOMP
	 DGT    = 1.0 + DGT
c---End Modified by ejo on 7/19/96
	 DGI(I) = 1.0/DG(I)                                             
	 STD(I) = ST(I)*DGT                                             
	 EDD(I) = DGT/DG(I)                                             
	 DO 20 J = 1,ND                                                 
	    DO 10 K = 1,NC                                              
	       BEDS(I,J,K) = (EDS(I) + D(I)*EDP(I))*EDD(I)*BR(J,K)      
   10       CONTINUE                                                    
   20    CONTINUE                                                       
	 DO 40 J = 1,ND                                                 
	    DO 30 K = 1,NC                                              
	       BEDP(I,J,K) = EDP(I)*(1.0 - D(I))*EDD(I)*BR(J,K)         
   30       CONTINUE                                                    
   40    CONTINUE                                                       
   50 CONTINUE                                                          

      IF (DEBUGM .EQ. 1) THEN
	OPEN(3,FILE='psdm.c2')
	I = 1
	WRITE(3,*) 'I =', I
	WRITE(3,*) 'DGT =', DGT
	WRITE(3,*) 'DGI(I) =', DGI(I)
	WRITE(3,*) 'STD(I) =', STD(I)
	WRITE(3,*) 'D(I) =', D(I)
	WRITE(3,*) 'EDD(I) =', EDD(I)
	WRITE(3,*) 'EDS(I) =', EDS(I)
	WRITE(3,*) 'EDP(I) =', EDP(I)

	DO 1021 J=1,ND
	  WRITE(3,*) (BR(J,K), K=1, NC)
	  WRITE(3,*) ' '
1021    CONTINUE

	DO 1011 I=1,ND
	  WRITE(3,*) (BEDS(1,I,J), J=1, NC)
	  WRITE(3,*) ' '
1011    CONTINUE
	DO 1012 I=1,ND
	  WRITE(3,*) (BEDP(1,I,J), J=1, NC)
	  WRITE(3,*) ' '
1012    CONTINUE
	CLOSE(3)
      END IF

      RETURN                                                            
      END                                                               
C
C
C                         FUNCTION CINF(I,T)
C                                                                       
C   **************************************************************      
C   * This function calculates the influent concentration to the *      
C   * column for each component at each time interval T.  If no  *      
C   * varying influent data is given this subroutine is ignored. *      
C   **************************************************************      
C                                                                       
C   Description of Inputs/Outputs/Return:
C   =====================================
C
C   *I*   I       Index of component
C   *I*   T       Time at which to calculate influent conc. (dim'less)
C   *R*   CINF    Calculated influent concentration (normalized, C/C0)
C

      DOUBLE PRECISION FUNCTION CINF(I,T)
      IMPLICIT NONE                              
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
C      INTEGER NCOMPI
C      PARAMETER (NCOMPI=6)
      INTEGER*2 BEDNUM
      INTEGER I,J
      INTEGER*2 MC,NC,NCOMP,N1,NIN
      DOUBLE PRECISION DGT
C      DOUBLE PRECISION CIN(NCOMPI,400),TIN(400),T
      DOUBLE PRECISION CIN(MXCOMP,400),TIN(400),T

      COMMON /BLOCKD/ CIN,TIN
      COMMON /BLOCKC/ MC,NC,NCOMP,N1,DGT,NIN
      COMMON /BLOCKH/ BEDNUM 

      IF (T .LE. TIN(1) ) THEN                                          
	 CINF = CIN(I,1)
      ELSE IF (T .GE. TIN(NIN) ) THEN                                   
	 CINF = CIN(I,NIN)                                              
      ELSE                                                              
	 J = 1                                                          
   10    J = J + 1                                                      
	    IF(T .GE. TIN(J-1) .AND. T .LE. TIN(J) ) THEN               
	       CINF = CIN(I,J-1) + (CIN(I,J)-CIN(I,J-1))*(T-TIN(J-1))/  
     +            (TIN(J)-TIN(J-1))                                     
	    ELSE IF (J .LT. NIN ) THEN                                  
	       GO TO 10                                                 
	    ENDIF                                                       
      ENDIF                                                             
      RETURN                                                            
      END                                                               
C                                                                       
C                    --------------------------                         
C                    I END OF SUBROUTINE CINF I                         
C                    --------------------------                         

C **********************************************************************                                                                       
		 SUBROUTINE PEDERV ( N,T,Y,PD,N0 )                      
C **********************************************************************                                                                       
c      IMPLICIT NONE
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
c      COMMON/BLOCKC/MC,NC,NCOMP,N1,DGT,NIN
      RETURN                                                            
      END                                                               

C--------------------------------------------------------------------
C    Subroutine CONSTANT - Provide Collocation Constants
C--------------------------------------------------------------------
      SUBROUTINE CONSTNT(N1,N2,AZ1,BR1,WR1,NFLAGO)
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
Cc**** Change Hokanson 2/8/97
Cc      PARAMETER (MCI=18,NCI=6)
C      PARAMETER (MCI=18,NCI=18)
Cc**** End Change Hokanson 2/8/97    

ccccccccccccccccccccccccccccc get rid of implicit double precision!
ccccccccccccccccccccccccccccc replace with implicit none

      INTEGER*2 N1
      INTEGER*2 N2
C      DOUBLE PRECISION AZ1(MCI,MCI),BR1(NCI,NCI),WR1(NCI)
      DOUBLE PRECISION AZ1(MAXMC,MAXMC),BR1(MAXNC,MAXNC),WR1(MAXNC)
      INTEGER*2 NFLAGO

      DOUBLE PRECISION R(18), Z(18), QI(18,18), RR(18)
      COMMON /QMTRXS/ Q(18,18), C(18,18), D(18,18), F(18)
      COMMON /OCCOEF/ AR(18,18), BR(18,18), AZ(18,18), BZ(18,18),
     +                WZ(18), WR(18)

C---- local variables
      INTEGER NR,NZ
      INTEGER NFLAG

c---- Debug variables
      INTEGER*2 DEBUGM
      DOUBLE PRECISION LAST_T
      COMMON /DEBUG/ LAST_T, DEBUGM

C     For Spherical coordinates:    NGEOR = 3
C                                   NOR =0
C                                   N1R = 1
C                                   ALFAR = 1.D0
C                                   BETAR = 0.5D0
C
C     For Cylindrical coordinates:  NGEOR = 2
C                                   NOR =0
C                                   N1R = 1
C                                   ALFAR = 1.D0
C                                   BETAR = 0.0D0
      DATA NGEOR/3/
     +    ,  N0R/0/,  N1R/1/,  ALFAR/1.0D0/,  BETAR/ 0.5D0/            
     +    ,  N0Z/1/,  N1Z/1/,  ALFAZ/0.0D0/,  BETAZ/ 0.0D0/            
      NR = N1
      NZ = N2
      NFLAGO = 0
      NFLAG = NFLAGO
									
      IF (DEBUGM .EQ. 1) THEN
	WRITE(4,*) 'N1 =', N1
	WRITE(4,*) 'N2 =', N2
	WRITE(4,*) 'NFLAG =', NFLAG
      END IF

      CALL DROOT(NR-1,N0R,N1R,ALFAR,BETAR,RR)                           
      DO 1 I = 1, NR                                                    
   1    R(I) = DSQRT(RR(I))                                             
      CALL DSPOLY(NGEOR, NR, NR, R)                                     
      CALL DLINRG( NR,  Q,18, QI,18,NFLAG)                                    
      IF (NFLAG.NE.0) then
	GOTO 9997
	endif
      
C     CALL DMRRRR(  1, NR,  F, 1, NR, NR, QI,18,  1, NR,WR,1,NFLAG)         
      CALL DMURRV( NR, NR, QI,18, NR, F, 2, NR, WR,NFLAG)                     
      IF (NFLAG.NE.0) then
	GOTO 9997
	endif
      
      CALL DMRRRR( NR, NR,  C,18, NR, NR, QI,18, NR, NR,AR,18,NFLAG)         
      IF (NFLAG.NE.0) then
	GOTO 9997
	endif
      
      CALL DMRRRR( NR, NR,  D,18, NR, NR, QI,18,NR,NR,BR,18,NFLAG)
      IF (NFLAG.NE.0) then
	GOTO 9997
	endif
	  
									
      CALL DROOT(NZ-2,N0Z,N1Z,ALFAZ,BETAZ, Z)                           
      CALL DUPOLY( NZ, NZ,  Z)                                          
      CALL DLINRG( NZ,  Q,18, QI,18,NFLAG)                                    
      IF (NFLAG.NE.0) then
	GOTO 9997
	endif
C     CALL DMRRRR(  1, NZ,  F, 1, NZ, NZ, QI,18,  1, NZ, WZ,1,NFLAG)         
      IF (NFLAG.NE.0) GOTO 9997
      CALL DMURRV( NZ, NZ, QI,18, NZ, F, 2, NZ, WZ,NFLAG)                     
      IF (NFLAG.NE.0) GOTO 9997
      CALL DMRRRR( NZ, NZ,  C,18, NZ, NZ, QI,18, NZ, NZ, AZ,18,NFLAG)         
      IF (NFLAG.NE.0) GOTO 9997
      CALL DMRRRR( NZ, NZ,  D,18, NZ, NZ, QI,18, NZ, NZ, BZ,18,NFLAG)         
      IF (NFLAG.NE.0) GOTO 9997

      DO 1515 I=1,NZ
	DO 1517 J=1,NZ
	  AZ1(I,J)=AZ(I,J)
1517   CONTINUE
1515  CONTINUE
      DO 1516 I=1,NR
	WR1(I)=WR(I)
	DO 1516 J=1,NR
	 BR1(I,J)=BR(I,J)
1516  CONTINUE

      IF (DEBUGM .EQ. 1) THEN
	WRITE(4,*) 'N1 =', N1
	WRITE(4,*) 'N2 =', N2
	WRITE(4,*) 'NFLAG =', NFLAG

	WRITE(4,*) 'AZ:'
	DO 2101 I=1,NZ
	  WRITE(4,*) (AZ(I,J), J=1, NZ)
2101    CONTINUE
	WRITE(4,*) 'BR:'
	DO 2102 I=1,NR
	  WRITE(4,*) (BR(I,J), J=1, NR)
2102    CONTINUE
	WRITE(4,*) 'WR:'
	DO 2103 I=1,NR
	  WRITE(4,*) WR(I)
2103    CONTINUE
      END IF


9997  NFLAGO = NFLAG

      if (DEBUGM .EQ. 1) then
        WRITE(4,*) 'NFLAGO = ', NFLAGO
        WRITE(4,*) 'NFLAG = ', NFLAG
      end if

      return
      END

C **********************************************************************
      SUBROUTINE DROOT(N,N0,N1,AL,BE,ROOT)                              
C **********************************************************************
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)                               
ccccccccccccccccccccccccccccc get rid of implicit double precision!
ccccccccccccccccccccccccccccc replace with implicit none

      DIMENSION D1(18), D2(18), ROOT(18)                        
      DATA ZERO/0.0D0/, ONE/1.0D0/, TWO/2.0D0/, TRE/3.0D0/              
     +     , DELTA/0.00010D0/, ERROR/0.10D-09/                          
      AB = AL + BE                                                      
      AD = BE - AL                                                      
      AP = BE * AL                                                      
      D1(1) = (AD / (AB + TWO) + ONE) / TWO                             
      D2(1) = ZERO                                                      
      IF(N.LT.2) GO TO 15                                               
      DO 10 I = 2, N                                                    
      Z1 = I - 1                                                        
      Z  = AB + TWO * Z1                                                
      D1(I) = (AB * AD / Z / (Z + TWO) + ONE) / TWO                     
      IF (I.NE.2) GO TO 11                                              
      D2(I) = (AB + AP + Z1) / Z / Z / (Z + ONE)                        
      GO TO 10                                                          
11      Z = Z * Z                                                       
	Y = Z1 * (AB + Z1)                                              
	Y = Y * (AP + Y)                                                
	D2(I) = Y / Z / (Z - ONE)                                       
10      CONTINUE                                                        
15      X = ZERO                                                        
      DO 20 I = 1, N                                                    
25      XD = ZERO                                                       
      XN = ONE                                                          
      XD1 = ZERO                                                        
      XN1 = ZERO                                                        
      DO 30 J = 1, N                                                    
	XP = (D1(J) - X) * XN - D2(J) * XD                              
	XP1 = (D1(J) - X) * XN1 - D2(J) * XD1 - XN                      
	XD = XN                                                         
	XD1 = XN1                                                       
	XN = XP                                                         
30      XN1 = XP1                                                       
      ZC = ONE                                                          
      Z = XN / XN1                                                      
      IF(I.EQ.1) GO TO 21                                               
      DO 22 J = 2, I                                                    
22      ZC = ZC - Z / (X - ROOT(J-1))                                   
21      Z = Z / ZC                                                      
      X = X - Z                                                         
      IF (DABS(Z).GT.ERROR) GO TO 25                                    
      ROOT(I) = X                                                       
20      X = X + DELTA                                                   
      NT = N + N0 + N1                                                  
      IF (N0.EQ.0)  GO TO 35                                            
      DO 31 I = 1,N                                                     
      J = N + 1-I                                                       
31      ROOT(J+1) = ROOT(J)                                             
      ROOT(1) = ZERO                                                    
35      IF(N1.EQ.1) ROOT(NT) = ONE                                      
      RETURN                                                            
      END                                                               
      
C **********************************************************************
      SUBROUTINE DUPOLY( ND, NC,  X)                                    
C **********************************************************************
      IMPLICIT DOUBLE PRECISION (A-H, O-Z)                              
ccccccccccccccccccccccccccccc get rid of implicit double precision!
ccccccccccccccccccccccccccccc replace with implicit none

      DIMENSION X(18), Y(36)                                            
      COMMON/QMTRXS/ Q(18,18), C(18,18), D(18,18), F(18)                
      DATA ZERO/0.0D0/, ONE/1.0D0/, TWO/2.0D0/, FOUR/4.0D0/             
      DO 10 J = 1, ND                                                   
      Y(1) = ONE                                                        
      DO 11 I = 2, NC                                                   
11      Y(I) = Y(I-1) * (TWO * X(J) - ONE)                              
      Q(J,1) = Y(1)                                                     
      Q(J,2) = Y(2)                                                     
      C(J,1) = ZERO                                                     
      C(J,2) = TWO                                                      
      D(J,1) = ZERO                                                     
      D(J,2) = ZERO                                                     
      IF(MOD(J,2).EQ.0) THEN                                            
      F(J)   = ZERO                                                     
      ELSE                                                              
      F(J)   = ONE / DFLOAT(J)                                          
      ENDIF                                                             
      DO 10 I = 3, NC                                                   
	Q(J,I) = Y(I)                                                   
	C(J,I) = TWO * (I - ONE) * Y(I-1)                               
10      D(J,I) = FOUR * (I - ONE) * (I - TWO) * Y(I-2)                  
      RETURN                                                            
      END                                                               

C **********************************************************************
      SUBROUTINE DSPOLY(IA,ND,NC,X)                                     
C **********************************************************************
      IMPLICIT DOUBLE PRECISION (A-H, O-Z)                              
ccccccccccccccccccccccccccccc get rid of implicit double precision!
ccccccccccccccccccccccccccccc replace with implicit none

      DIMENSION X(18), Y(36)                                            
      COMMON/QMTRXS/ Q(18,18), C(18,18), D(18,18), F(18)                
      DATA ZERO/0.0D0/, ONE/1.0D0/, TWO/2.0D0/, FOUR/4.0D0/             
      DO 10 J = 1, ND                                                   
      Y(1) = X(J)                                                       
      DO 11 I = 2, 2 * NC                                               
11      Y(I) = Y(I-1) * X(J)                                            
      Q(J,1) = ONE                                                      
      Q(J,2) = Y(2)                                                     
      C(J,1) = ZERO                                                     
      C(J,2) = TWO * Y(1)                                               
      D(J,1) = ZERO                                                     
      D(J,2) = TWO * IA                                                 
      F(J)   = ONE / DFLOAT(2 * J - 2 + IA)                             
      DO 10 I = 3, NC                                                   
	Q(J,I) = Y(2*I-2)                                               
	C(J,I) = (TWO * I - TWO) * Y(2*I-3)                             
10      D(J,I) = (TWO * I - TWO) * (TWO * I - FOUR + IA) * Y(2*I-4)     
      RETURN                                                            
      END
C

