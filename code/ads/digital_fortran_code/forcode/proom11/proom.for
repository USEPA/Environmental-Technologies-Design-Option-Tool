CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
CC
CC  Program Name:       PSDM for DLL
CC  Author:             Michigan Tech University - 1994
CC  Intended Platform:  Compiled with Microsoft FORTRAN and linked
CC                      to the Visual Basic code of the Adsorption
CC                      Simulation Software.
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
CC  12/07/1997: Eric Oman
CC  - Modified to incorporate backwashing; refer to Gary Friedman's
CC    thesis, pp. 74-76 and appendices 12 and 28.
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
CC                    Array Size: (5)
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



c      SUBROUTINE PSDM(Numb,Chemicals,Ads_Prop,C_Prop,
c     &                T,CPVB,NITP,TT,NXX,MXX,
c     &                NinI,TinI,CinI,N_PW,NumBed,NFLAG,
c     &                VARS1,VARS2,ISDBUG)
c      SUBROUTINE PSDM(Numb,Chemicals,Ads_Prop,C_Prop,
c     &                T,CPVB,NITP,TT,NXX,MXX,
c     &                NinI,TinI,CinI,NumBed,NFLAG,
c     &                VARS1,VARS2,ISDBUG,
c     &                TELL_PSDM_SPECIAL_OUTPUT)
C      SUBROUTINE PSDM(Numb,Chemicals,Ads_Prop,C_Prop,
C     &  T,CPVB,NITP,TT,NXX,MXX,
C     &  NinI,TinI,CinI,NumBed,NFLAG,
C     &  VARS1,VARS2,ISDBUG,
C     &  TELL_PSDM_SPECIAL_OUTPUT,NB,TBACK,
C     &  in_IS_IN_ROOM,in_ROOM_VOL,in_ROOM_FLOWRATE,in_ROOM_C0,
C     &  in_ROOM_EMIT,in_FN_MASSBAL_OUT,in_FN_CR_OUT,in_FN_CB_OUT)
      SUBROUTINE PSDM(Numb,Chemicals,Ads_Prop,C_Prop,
     &  T,CPVB,NITP,TT,NXX,MXX,
     &  NinI,TinI,CinI,NumBed,NFLAG,
     &  VARS1,VARS2,ISDBUG,
     &  TELL_PSDM_SPECIAL_OUTPUT,NB,TBACK,
     &  in_IS_IN_ROOM,in_ROOM_VOL,in_ROOM_FLOWRATE,in_ROOM_C0,
     &  in_ROOM_EMIT,
     &  in_FN_MASSBAL_OUT,in_FN_CR_OUT,in_FN_CB_OUT,
     &  in_INITIAL_ROOM_CONC)
c     &  in_RXN_RATE_CONSTANT,in_RXN_PRODUCT,in_RXN_RATIO,
      IMPLICIT NONE
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'

c---- Constants
      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE

c**** Change Hokanson 2/8/97
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=6,MAXPTS=400,MAXDE=750)
      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=6000,MAXDE=750)
c**** End Change Hokanson 2/8/97
c new maximums:
      INTEGER MXTBACK
      PARAMETER (MXTBACK=400)

      INTEGER*2 NSTEPS
c      PARAMETER (NSTEPS=400)
      PARAMETER (NSTEPS=6000)

cC**** Define Workspace **********************************************
cc     The HUGE argument should not be needed since the DLL is
cc     compiled using the /AH option.
c      DOUBLE PRECISION PW[ALLOCATABLE,HUGE](:)

C      
C     N_PW IS ADDED TO THIS PARAMETER STATEMENT TO ALLOCATE THE
C     WORKSPACE FOR PSDM AND DGEAR.  THIS VALUE CORRESPONDS TO
C     THE CAPABILITY FOR A MAXIMUM NC = 6, MC = 18, AND 6 COMPONENTS.
C     THIS WAS PREVIOUSLY DONE IN VISUAL BASIC. 
      
      INTEGER*4 N_PW
      PARAMETER(N_PW=1336720)
      DOUBLE PRECISION PW(N_PW)

c---- Input variables
      INTEGER*2 Numb
      DOUBLE PRECISION Chemicals(MXCOMP,16)
      DOUBLE PRECISION Ads_Prop(4)
      DOUBLE PRECISION C_Prop(3)
      DOUBLE PRECISION T(MAXPTS,2)
      DOUBLE PRECISION CPVB(MXCOMP,MAXPTS)
      INTEGER*2 NITP
      DOUBLE PRECISION TT(5)
      INTEGER*2 NXX,MXX,NinI
      DOUBLE PRECISION TinI(NinI),CinI(MXCOMP,NinI)
c      INTEGER*4 N_PW
      INTEGER*2 NumBed
      INTEGER*2 NFLAG
      DOUBLE PRECISION VARS1(15)
      DOUBLE PRECISION VARS2(MXCOMP,19)
      INTEGER*4 ISDBUG
      INTEGER*4 TELL_PSDM_SPECIAL_OUTPUT
c new parameters to psdm() subroutine:
      INTEGER*4 NB
      DOUBLE PRECISION TBACK(MXTBACK)
c new parameters to psdm() subroutine:
      INTEGER in_IS_IN_ROOM
      DOUBLE PRECISION in_ROOM_VOL
      DOUBLE PRECISION in_ROOM_FLOWRATE
      DOUBLE PRECISION in_ROOM_C0(1:MXCOMP)
      DOUBLE PRECISION in_ROOM_EMIT(1:MXCOMP)
c      DOUBLE PRECISION in_RXN_RATE_CONSTANT(1:MXCOMP)
c      DOUBLE PRECISION in_RXN_PRODUCT(1:MXCOMP)
c      DOUBLE PRECISION in_RXN_RATIO(1:MXCOMP)
      CHARACTER*100 in_FN_MASSBAL_OUT
      CHARACTER*100 in_FN_CR_OUT
      CHARACTER*100 in_FN_CB_OUT
      DOUBLE PRECISION in_INITIAL_ROOM_CONC(1:MXCOMP)
      INTEGER IS_IN_ROOM
      DOUBLE PRECISION ROOM_VOL
      DOUBLE PRECISION ROOM_FLOWRATE
      DOUBLE PRECISION ROOM_C0(1:MXCOMP)
      DOUBLE PRECISION ROOM_EMIT(1:MXCOMP)
      CHARACTER*100 FN_MASSBAL_OUT
      CHARACTER*100 FN_CR_OUT
      CHARACTER*100 FN_CB_OUT
c
c---- REACTION-RELATED STUFF.
c
c      DOUBLE PRECISION RXN_RATE_CONSTANT(1:MXCOMP)
c      DOUBLE PRECISION RXN_PRODUCT(1:MXCOMP)
c      DOUBLE PRECISION RXN_RATIO(1:MXCOMP)
c      COMMON /RXN1/ RXN_RATE_CONSTANT,RXN_PRODUCT,RXN_RATIO

c---- Storage of input variables
      INTEGER*2 NCOMP
      DOUBLE PRECISION XWT(1:MXCOMP)
      DOUBLE PRECISION CBO(MXCOMP),VB(MXCOMP),XK(MXCOMP),
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
      DOUBLE PRECISION CR(MXCOMP,MAXPTS)
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

c---- NEW LOCAL VARIABLES (12/7/97):
      INTEGER*4 NA
      INTEGER*4 XNB
      DOUBLE PRECISION XTBACK(MXTBACK)
      DOUBLE PRECISION WZ1(MAXNC)
      DOUBLE PRECISION WZ(MAXNC)

c new parameters to psdm() subroutine:
      COMMON /AMWAY1/ IS_IN_ROOM,ROOM_VOL,ROOM_FLOWRATE,ROOM_C0,
     &  ROOM_EMIT,FN_MASSBAL_OUT
      COMMON /AMWAY2/ XWT,FLRT

c---- NEW LOCAL VARIABLES (6/24/98):
      DOUBLE PRECISION MASSBAL_(1:MXCOMP,1:7)
      DOUBLE PRECISION CR_(1:MXCOMP)
      DOUBLE PRECISION CB_(1:MXCOMP)
      INTEGER CRIDX_(1:MXCOMP)

C---- NEW LOCAL VARIABLES (9/16/98):
      DOUBLE PRECISION INITIAL_ROOM_CONC(1:MXCOMP)

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

c      ALLOCATE (PW(N_PW),STAT=error)
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
      IF (NIN .EQ. 0) GO TO 811
      DO 1 J = 1,NIN                                                    
C       READ(4,*) TIN(J), (CIN(I,J), I = 1,NCOMP)                      
	TIN(J) = TinI(J)
	DO 2 I=1,NCOMP
	  CIN(I,J) = CinI(I,J)
2       CONTINUE
1     CONTINUE

c---- Input backwashing variables
811   XNB = NB
      DO I=1,XNB
        XTBACK(I) = TBACK(I)
      ENDDO

c---- Input filter-in-room related variables
      IS_IN_ROOM = in_IS_IN_ROOM
      ROOM_VOL = in_ROOM_VOL
      ROOM_FLOWRATE = in_ROOM_FLOWRATE
      DO I=1,NCOMP
        ROOM_C0(I) = in_ROOM_C0(I)
        ROOM_EMIT(I) = in_ROOM_EMIT(I)
c        RXN_RATE_CONSTANT(I) = in_RXN_RATE_CONSTANT(I)
c        RXN_PRODUCT(I) = in_RXN_PRODUCT(I)
c        RXN_RATIO(I) = in_RXN_RATIO(I)
      ENDDO
      FN_MASSBAL_OUT = in_FN_MASSBAL_OUT
      FN_CR_OUT = in_FN_CR_OUT
      FN_CB_OUT = in_FN_CB_OUT
							  
C
C                  CALCULATIONS
C

c---- Calculate number of equations
812   NEQ = (MC*(NC + 1) - 1)*NCOMP
      IF (IS_IN_ROOM.EQ.1) THEN
        NEQ = NEQ + NCOMP
      ENDIF

c---- Convert influent concentrations from ug/L ---> umol/L
      DO 212 I = 1, NCOMP
	CBO(I) = CBO(I)/XWT(I)
	DO 211 J=1, NIN
	  CIN(I,J) = CIN(I,J)/XWT(I)
211     CONTINUE
212   CONTINUE                                                          

C
C---- KEEP ROOM INITIAL CONCENTRATIONS IN UNITS OF ug/L.
C
      DO I=1, NCOMP
        INITIAL_ROOM_CONC(I) = in_INITIAL_ROOM_CONC(I)
      ENDDO

c---- Calculate various bed parameters
      AREA = 3.141592654D0*DIA*DIA/4.0D0
      BEDVOL = L*AREA                                                   
      EBED = 1.0D0 - WT/(BEDVOL*RHOP)                                   
      EBCT = BEDVOL/FLRT                                                
c      EBCT = BEDVOL/(FLRT*1.0d-6)
      TAU = BEDVOL*EBED*60.0D0/FLRT                                     

c---- Calculate collocation constants
      MCA = MC
      NCA = NC
      CALL CONSTNT(NCA,MCA,AZ1,BR1,WR1,WZ1,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      DO 9 I=1,MC
	  WZ(I)=WZ1(I)
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
C      IF (IS_IN_ROOM) THEN
C        N = N + NCOMP
C      ENDIF

c---- Initialize dependent variables
      DO I = 1, N
        Y0(I) = 0.0
      ENDDO
      IF (IS_IN_ROOM.EQ.1) THEN
C
C...... INITIALIZE ROOM CONCENTRATION FOR EACH CHEMICAL.
C
        DO I=1, NCOMP
C          Y0(N+I) = 0.0D0
          Y0(N+I) = INITIAL_ROOM_CONC(I)
        ENDDO
        N = N + NCOMP
      ENDIF

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

      DO I=1, XNB
        XTBACK(I) = XTBACK(I)*TCONV
      ENDDO

C
C------ IF PSDM-IN-ROOM MODEL, OUTPUT CONCENTRATIONS ON THE FLY.
C
C////////// Note change by Eric Oman on 10/21/99 below.
C      IF (IS_IN_ROOM.EQ.1) THEN
      IF (1.EQ.1) THEN
      print *, 'FN_MASSBAL_OUT=', FN_MASSBAL_OUT
        OPEN(UNIT=19,FILE=FN_MASSBAL_OUT)
        WRITE(19,*) 'NOTE: ALL TERMS IN UNITS OF UG/S.'
        OPEN(UNIT=21,FILE=FN_CR_OUT)
        OPEN(UNIT=22,FILE=FN_CB_OUT)
      ENDIF
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
      NA = 1

C
C------ SET UP VARIOUS PSDM-IN-ROOM PARAMETERS.
C
      IF (1.EQ.1) THEN
        EPOR_ = EPOR
        RHOP_ = RHOP
        EBED_ = EBED
        TAU_ = TAU
        DO I=1,NCOMP
          QE_(I) = QE(I)
        ENDDO
      ENDIF

   70 ITP = ITP + 1                                                     

      CALL DGEAR (N,T0,H0,Y0,TOUT,EPS,MF,INDEX,NFLAG,PW,N_PW)
      WRITE (*,'(1X,''PERCENT COMPLETE = '',F7.2,''%'')') 
     &      (100.0D0*TOUT)/TTOL

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

C
C------ STORE THE CURRENT TIME POINT.   
C
C////////// Note change by Eric Oman on 10/21/99 below.
Cxaxa
      WRITE(21,*) TOUT/TCONV,(Y0(N1*NCOMP+I),I=1,NCOMP)
c      WRITE(22,*) TOUT/TCONV,(Y0(N1*I)*CBO(I),I=1,NCOMP)
      WRITE(22,*) TOUT/TCONV,(Y0(N1*I)*CBO(I)*XWT(I),I=1,NCOMP)
C      IF (IS_IN_ROOM.EQ.1) THEN
Cc              print *, ncomp
Cc        WRITE(21,1035) T/TCONV,(Y0(N1*NCOMP+I),I=1,NCOMP)
Cc        WRITE(22,1035) T/TCONV,(Y0(N1*I),I=1,NCOMP)
C        WRITE(21,*) TOUT/TCONV,(Y0(N1*NCOMP+I),I=1,NCOMP)
C        WRITE(22,*) TOUT/TCONV,(Y0(N1*I),I=1,NCOMP)
C      ELSE
C        DO I = 1,NCOMP                                                 
C	    CP(I,ITP) = Y0(N1*I)
CC        CR(I,ITP) = Y0(N1*NCOMP + I)
C        ENDDO
C      ENDIF
C////////// Note change by Eric Oman on 10/21/99 ends.
      IF (IS_IN_ROOM.EQ.1) THEN
            DO I=1,NCOMP
C              CRIDX_(I) = N1*NCOMP + I
C              CR_(I) = Y0(CRIDX_(I))
              CR_(I) = Y0(N1*NCOMP+I)
CC---------- (CB_,UG/L) = (Y0,DIM'LESS)*(CBO,UMOL/L)*(XWT,UG/UMOL)
C              CB_(I) = Y0(N1*I)*CBO(I)*XWT(I)
C---------- (CB_,UG/L) = (Y0,DIM'LESS)*(CBO,UMOL/L)*(XWT,UG/UMOL)
              CB_(I) = Y0(N1*I)*CBO(I)*XWT(I)
              IF (CB_(I).LT.0.0D0) THEN
                CB_(I) = 0.0D0
              ENDIF
C------- (MASSBAL_,UG/S) = (ROOM_FLOWRATE,L/S)*(ROOM_C0,UG/L)
              MASSBAL_(I,1) = (ROOM_FLOWRATE/1000.0D0)*ROOM_C0(I)
C------- (MASSBAL_,UG/S) = (ROOM_FLOWRATE,L/S)*(CR_,UG/L)
              MASSBAL_(I,2) = (ROOM_FLOWRATE/1000.0D0)*CR_(I)
C------- (MASSBAL_,UG/S) = (ROOM_EMIT,UG/S)
              MASSBAL_(I,3) = ROOM_EMIT(I)
C------- (MASSBAL_,UG/S) = (FLRT,L/S)*(CR_,UG/L)
              MASSBAL_(I,4) = (FLRT/60.0D0/1000.0D0)*CR_(I)
C------- (MASSBAL_,UG/S) = (FLRT,L/S)*(CB_,UG/L)
              MASSBAL_(I,5) = (FLRT/60.0D0/1000.0D0)*CB_(I)
              MASSBAL_(I,6) = 
     &  MASSBAL_(I,1) + MASSBAL_(I,3) + MASSBAL_(I,5)
              MASSBAL_(I,7) = MASSBAL_(I,2) + MASSBAL_(I,4)
              MASSBAL_(I,8) = MASSBAL_(I,6) - MASSBAL_(I,7)
            ENDDO
            WRITE (19,'(7(G20.10))') TOUT/TCONV,
     &        MASSBAL_(1,1),
     &        MASSBAL_(1,2),
     &        MASSBAL_(1,3),
     &        MASSBAL_(1,4),
     &        MASSBAL_(1,5),
     &        MASSBAL_(1,7)
      ENDIF

      TP(ITP) = TOUT                                                    
      DOUT = TOUT/TCONV                                                 
      IF ( ITP .LT. NSTEPS ) THEN                                       
        IF ( TOUT .LT. TTOL ) THEN
          IF (1.NE.1) THEN
            IF ( (XNB.NE.0) .AND. (TOUT.GE.XTBACK(NA)) ) THEN 
c               WRITE(7,114) 
               CALL WASH( Y0,WZ )   
               T0 = TOUT
C               H0 = H01 
               INDEX = 1
               IF ( NA .EQ. XNB ) THEN   
                  XNB = 0
               ELSE 
                  NA = NA + 1   
               ENDIF
            ENDIF   
          ENDIF
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
c	DEALLOCATE(PW,STAT=error)
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

      IF (TELL_PSDM_SPECIAL_OUTPUT.eq.1) then
C
C    Conversion of bed length,weight,
C    tau, ST, EDS, and EDP back to the entire bed value instead of per
C    each axial element.
C
        EBCT = EBCT * DBLE(NUMBED)
        L = L * DBLE(NUMBED)
        WT = WT * DBLE(NUMBED)
        TAU = TAU * DBLE(NUMBED)
        DO I = 1,NCOMP
	    ST(I) = ST(I) * DBLE(NUMBED)
	    EDP(I) = EDP(I) * DBLE(NUMBED)
	    EDS(I) = EDS(I) * DBLE(NUMBED)
        ENDDO 

c        open (5,file='psdm.out')
        open (5,file=FN_OUT_MAIN)
        WRITE(5,1005) NCOMP,NC,MC,NUMBED,NEQ,NFLAG,RAD,RHOP,EPOR,
     1  	L,DIA,WT,FLRT,EBED,TAU,EBCT
        WRITE(5,*) '(*) Note K units are (umol/g)*(L/umol)^(1/n)'
        DO I=1,NCOMP
          WRITE(5,1011) '?',VB(I),XWT(I),CBO(I),XK(I),XN(I),
     &      KF(I),DS(I),DP(I),TOR(I),D(I),
     &      ST(I),DGS(I),BIS(I),EDS(I),EDP(I),DGP(I),BIP(I)
        ENDDO

        IF (IS_IN_ROOM.EQ.1) THEN
C-------- DO NOTHING.    
        ELSE
          DO J = 1,ITP
            WRITE(5,1035) T(J,1),(CP(I,J),I=1,NCOMP)
C ,(CR(I,J),I=1,NCOMP)
          ENDDO
1035      FORMAT(1X,E18.9,1X,6(1X,E18.9),6(1X,E18.9))
        ENDIF

 1005   FORMAT(////
     &   'NUMBER OF COMPONENTS, NCOMP............... = ',I15/
     &   'NUMBER OF RADIAL COLLOCATION POINTS, NC... = ',I15/
     &   'NUMBER OF AXIAL COLLOCATION POINTS, MC.... = ',I15/
     &   'CURRENT AXIAL ELEMENT, NUMBED............. = ',I15/
     &   'TOTAL NO. OF DIFFERENTIAL EQUATIONS, NEQ.. = ',I15/
     &   'ERROR FLAG FROM DGEAR, NFLAG.............. = ',I15/
     &   'RADIUS OF ADSORBENT PARTICLE, RAD (CM) ... = ',G20.10/
     &   'APPARENT DENSITY OF ADSORBENT, RHOP (G/ML) = ',G20.10/
     &   'ADSORBENT PARTICLE POROSITY, EPOR......... = ',G20.10/
     &   'BED LENGTH, L, (CM)....................... = ',G20.10/
     &   'BED DIAMETER, DIA, (CM)................... = ',G20.10/
     &   'ADSORBENT WEIGHT, WT (G).................. = ',G20.10/
     &   'FLOWRATE, FLRT (ML/MIN)................... = ',G20.10/
     &   'BED POROSITY, EBED........................ = ',G20.10/
c     &   'SUPERFICIAL VELOCITY, VS (CM/MIN)......... = ',G20.10/
     &   'PACKED BED CONTACT TIME, TAU (SEC)........ = ',G20.10/
     &   'EMPTY BED CONTACT TIME, EBCT (MIN)........ = ',G20.10//)
c     &   'REYNOLDS NUMBER, RE....................... = ',G20.10/
c     &   'TEMPERATURE, TEMP (CELSIUS)............... = ',G20.10/
c     &   'PRESSURE, ATM (ATM)....................... = ',G20.10/
c     &   'DENSITY OF WATER, DW (G/ML)............... = ',G20.10/
c     &   'VISCOSITY OF WATER, VW (CP)............... = ',G20.10/
c     &   'DENSITY OF PURGE GAS, DPURG (G/ML)........ = ',G20.10/
c     &   'VISCOSITY OF PURGE GAS, VPURG (KG/M-SEC).. = ',G20.10/
c     &   'MOLECULAR WEIGHT OF PURGE GAS, GAS_MW..... = ',G20.10//)
 1011 FORMAT(/' PARAMETERS FOR  ',A30/    
     & 4X,'MOLAR VOLUME AT BOILING PT.,VB (ML/GMOL). = ',G20.10/
c     & 4X,'NORMAL BOILING POINT, NBP (KELVIN)....... = ',G20.10/
     & 4X,'MOLECULAR WEIGHT, XWT (DALTON)........... = ',G20.10/
     & 4X,'INITIAL BULK CONCENTRATION, CBO (UMOL/L). = ',G20.10/
     & 4X,'FREUNDLICH K, XK, (*).................... = ',G20.10/
     & 4X,'FREUNDLICH 1/N, XN....................... = ',G20.10/
c     & 4X,'DIFFUSIVITY, DIFFUS, (CM^2/SEC).......... = ',G20.10/
c     & 4X,'SCMIDT NUMBER, SC........................ = ',G20.10/
     & 4X,'FILM TRANSFER COEFF., KF, (CM/SEC)....... = ',G20.10/
     & 4X,'SURFACE DIFFUSION COEFF., DS, (CM^2/SEC). = ',G20.10/
     & 4X,'PORE DIFFUSION COEFF., DP, (CM^2/SEC).... = ',G20.10/
c     & 4X,'SURFACE TO PORE DIFFUSION RATIO, SPDFR... = ',G20.10/
     & 4X,'TORTUOSITY CONSTANT, TOR................. = ',G20.10/
     & 4X,'DS OVER DP, D............................ = ',G20.10/
     & 4X,'STANTON NUMBER, ST....................... = ',G20.10/
     & 4X,'SURFACE SOLUTE DIST. PARAMETER, DGS...... = ',G20.10/
     & 4X,'SURFACE BIOT NUMBER, BIS................. = ',G20.10/
     & 4X,'SURFACE DIFFUSION MODULUS, EDS........... = ',G20.10/
     & 4X,'PORE DIFFUSION MODULUS, EDP.............. = ',G20.10/
     & 4X,'PORE SOLUTE DIST. PARAMETER, DGP......... = ',G20.10/
     & 4X,'PORE BIOT NUMBER, BIP.................... = ',G20.10//)

        close (5)
      endif

      IF (DEBUGM .EQ. 1) THEN
	CLOSE(4)
	CLOSE(8)
      END IF

C////////// Note change by Eric Oman on 10/21/99 below.
c      IF (IS_IN_ROOM.EQ.1) THEN
      IF (1.EQ.1) THEN
C        WRITE (19,*) '-1    -1'
        CLOSE(19)
        WRITE (21,*) 'END_OF_DATA'
        WRITE (21,*) 'EXIT VALUE OF NFLAG ='
        WRITE (21,*) NFLAG
        WRITE (21,*) 'NOTE: CONCENTRATION UNITS ARE UG/L'
        CLOSE(21)
        WRITE (22,*) 'END_OF_DATA'
        WRITE (22,*) 'EXIT VALUE OF NFLAG ='
        WRITE (22,*) NFLAG
cXAXA
c        WRITE (22,*) 'NOTE: CONCENTRATION UNITS ARE UMOL/L'
        WRITE (22,*) 'NOTE: CONCENTRATION UNITS ARE UG/L'
        CLOSE(22)
      ENDIF

      RETURN
      END                                                               
C                                                                       
C **********************************************************************                                                                       
		       SUBROUTINE ORTHOG( N )                          
C **********************************************************************                                                                      
c      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

c---- Subroutine parameters
      INTEGER*4 N

c---- Constants
      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE

c**** Change Hokanson 2/8/97
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=6,MAXPTS=400,MAXDE=750)
      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=6000,MAXDE=750)
c**** End Change Hokanson 2/8/97

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

      INTEGER NCOMPI
      PARAMETER (NCOMPI=6)
      INTEGER*2 BEDNUM
      INTEGER I,J
      INTEGER*2 MC,NC,NCOMP,N1,NIN
      DOUBLE PRECISION DGT
      DOUBLE PRECISION CIN(NCOMPI,400),TIN(400),T

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
c      SUBROUTINE CONSTNT(N1,N2,AZ1,BR1,WR1,NFLAGO)
      SUBROUTINE CONSTNT(N1,N2,AZ1,BR1,WR1,WZ1,NFLAGO)

c**** Change Hokanson 2/8/97
c      PARAMETER (MCI=18,NCI=6)
      PARAMETER (MCI=18,NCI=18)
c**** End Change Hokanson 2/8/97    

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
ccccccccccccccccccccccccccccc get rid of implicit double precision!
ccccccccccccccccccccccccccccc replace with implicit none

      INTEGER*2 N1
      INTEGER*2 N2
      DOUBLE PRECISION AZ1(MCI,MCI),BR1(NCI,NCI),WR1(NCI)
      DOUBLE PRECISION WZ1(NCI)
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
	WRITE(4,*) 'NC = N1 =', N1
	WRITE(4,*) 'MC = N2 =', N2
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
	  WZ1(I)=WZ(I)
	  DO 1517 J=1,NZ
	    AZ1(I,J)=AZ(I,J)
1517    CONTINUE
1515  CONTINUE
      DO 1516 I=1,NR
	  WR1(I)=WR(I)
	  DO 1516 J=1,NR
	    BR1(I,J)=BR(I,J)
1516  CONTINUE

      IF (DEBUGM .EQ. 1) THEN
      open (unit=8,file='constnt.txt')
	WRITE(8,*) 'NC = N1 =', N1
	WRITE(8,*) 'MC = N2 =', N2
	WRITE(8,*) 'NFLAG =', NFLAG

	WRITE(8,*) 'AZ:'
	DO 2101 I=1,NZ
	  WRITE(8,*) (AZ(I,J), J=1, NZ)
2101  CONTINUE
	WRITE(8,*) 'BR:'
	DO 2102 I=1,NR
	  WRITE(8,*) (BR(I,J), J=1, NR)
2102  CONTINUE
	WRITE(8,*) 'WR:'
	DO 2103 I=1,NR
	  WRITE(8,*) WR(I)
2103  CONTINUE
	WRITE(8,*) 'WZ:'
	DO 2104 I=1,NZ
	  WRITE(8,*) WZ(I)
2104  CONTINUE
      close (8)
      stop
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



                     SUBROUTINE WASH ( Y0,WZ )  
C   
C   *************************************************************** 
C   * This subroutine simulates a backwashing process by reset-   * 
C   * ting the concetrations in both the liquid and solid phases  * 
C   * by assuming a completely mixed adsorber during the process. * 
C   * The solid phase concentration gradients for each axial pos- * 
C   * ition in the bed (surface and pore) is set to the weighted  * 
C   * average using the axial weighting vector.  The liquid phase * 
C   * concentrations are set to an initial concentration of one.  * 
C   *************************************************************** 
C   

      INTEGER*2 MXCOMP,MAXMC,MAXNC,MAXPTS,MAXDE
      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=400,MAXDE=750)
c      PARAMETER (MXCOMP=6,MAXMC=18,MAXNC=18,MAXPTS=6000,MAXDE=750)
      DOUBLE PRECISION Y0(MAXDE)
      DOUBLE PRECISION WZ(MAXNC)
      DOUBLE PRECISION CPORE(MAXDE)
      COMMON /WASH1/ CPORE

c      DIMENSION Y0(1),WZ(1) 
c      COMMON/BLOCKF/CPORE(400)  
      COMMON MC,NC,NCOMP,N1,DGT,NIN,MND,MD,ND   

      STOP 'debug to make sure parameters passed properly!'

      DO 60 I = 1,NCOMP 
         II = (I - 1)*N1
         III = II + MND 
         IIII = III + MD
         DO 30 J = 1,ND 
            JJ = II + (J - 1)   
            QQ = 0.0
            QR = 0.0
            DO 10 K = 1,MC  
               KK = (K - 1)*ND + JJ 
               QQ = QQ + WZ(K)*Y0(KK + 1)   
               QR = QR + WZ(K)*CPORE(KK + 1)
   10       CONTINUE
            DO 20 K = 1,MC  
               KK = (K - 1)*ND + JJ 
               Y0(KK + 1) = QQ  
               CPORE(KK + 1) = QR   
   20       CONTINUE
   30    CONTINUE   
         QQ = 0.0   
         QR = 0.0   
         DO 40 K = 1,MC 
            QQ = QQ + WZ(K)*Y0(III + K) 
            QR = QR + WZ(K)*CPORE(III + K)  
   40    CONTINUE   
         DO 50 K = 1,MC 
            Y0(III+ K) = QQ 
            CPORE(III+ K) = QR  
   50    CONTINUE   
         DO 55 K = 2,MC 
            Y0(IIII + K) = 1.0  
   55    CONTINUE   
   60 CONTINUE  
      RETURN
      END   
C
C   
C                    -------------------------- 
C                    I END OF SUBROUTINE WASH I 
C                    -------------------------- 
