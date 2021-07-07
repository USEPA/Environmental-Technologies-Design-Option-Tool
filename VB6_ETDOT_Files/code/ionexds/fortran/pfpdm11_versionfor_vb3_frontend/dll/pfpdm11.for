C----- Ion Exchange Program
C----- Program for Microsoft FORTRAN
C----- PFPDM with Langmuir for Ion Exchange Equilibrium
C----- Michigan Tech University - 1994      
C----- Created 1/13/95 by D. Hokanson
C----- Uses DGEAR subroutine
C
C----- History:
C-------- Hokanson 12-Jan-01
C      Fixed problem in calculation of CPORE within
C      DIFFUN in that CTAVG was incorrectly calculated
C      for the case of variable influent.  This error
C      caused behavior that made the program produce
C      differing results as the number of axial 
C      elements (tanks) was changed, which was 
C      incorrect.  The program now produces the same
C      results when number of axial elements is the
C      only quantity varied as expected.  A previous
C      change made that was thought to help fix this
C      problem was removed from the code: breaking 
C      out of the program when all ions had reached 
C      1% of their influent concs.  The program now
C      runs to completion with no test for breaking out
C      in this manner (PFPDM11). 
C-------- Hokanson 11-Aug-00
C      Modified to pass in EPS and DH0 as input 
C      parameters (PFPDM09).
C-------- Hokanson 12-Aug-00
C      Modified to write C (mg/L) versus Time (min)
C      to output file named PDM10MG.OUT and
C      C (meq/L) versus Time (days) to output
C      file named PDM10MEQ.OUT and C/CT versus
C      Time in minutes in file named PDM10DIM.OUT
C      (PFPDM10).
C
C----------------------------------------------------
C------ List of Input Parameters --------------------
C----------------------------------------------------
C     Ads_Prop  : Adsorber Data                 Input
C                  Size 4
C                (1) -> Length (m)
C                (2) -> Diameter (m)
C                (3) -> Weight of GAC (kg)
C                (4) -> Inlet flowrate (m3/s)
C     Chemicals : Components Properties         Input
C                  Size Number x 9
C                (I,1) -> MW (g/mol)
C                (I,2) -> Initial conc. (mg/l)
C                (I,3) -> Separation Factor
C                         Alpha(I,Presaturant)
C                (I,4) -> kf (cm/s)
C                (I,5) -> Ds (cm2/s)
C                (I,6) -> Valence
C                (I,7) -> q/qT value of this constituent 
C                         at the initial condition (i.e. the
C                         amount of this constituent that
C                         presaturates the adsorbent)
C     Cin       : variable influent concentrations (mg/L) Input
C                 size Number x Nin
C     CT_AVERAGE : Sum of Time Variable Influent Concentrations
C                 Array of size Number
C     C_Prop    : Carbon Properties             Input
C                 Array of size 7
C                (1) -> Void Fraction of the particle 
C                (2) -> Apparent Density (g/cm3)
C                (3) -> Particle Radius(cm)
C                (4) -> Tortuosity
C                (5) -> Tor  Coeff for Tortuosity=f(t)
C                (6) -> Part Coeff for Tortuosity=f(t)
C                (7) -> Time parameter(mn) for Tortuosity=f(t)
C     Nin       : Number of influent conc. points Input 
C     MXXX      : number of Axial               Input
C                 Collocation Points
C     NXXX      : Number of Radial              Input
C                 Collocation Points
C     N_PW      : size of the working space     Input
C     Number    : Number of components          Input
C     Tin       : Time points for inf. conc.    Input
C                 Size Nin
C     TT        : Time parameters for DGEAR     Input   
C                 Size 5   
C ****** K is in (umol/g)x(l/umol)^(1/n) ************
C
C----------------------------------------------------
C------ List of Input Parameters --------------------
C----------------------------------------------------
C     CPVB        : Reduced brekthrouugh 
C                  concentrations(-)             Output
C                 Size: Number x 400
C     N_FLAG    : Error Flag                    Output
C     T         : Time points for CP (min)      Output
C
C------------------------------------------------
C----- Subroutine PFPDM10 = Main Subroutine
C------------------------------------------------
C---- begin modification hokanson 11-Aug-2000: PFPDM09
C      SUBROUTINE PFPDM08(Numb,Chemicals,Ads_Prop,C_Prop,
C     &                   T,CPVB,NITP,TT,NXX,MXX,
C     &                   NinI,TinI,CinI,CT_AVERAGE,N_PW,NFLAG)
      SUBROUTINE PFPDM11(Numb,Chemicals,Ads_Prop,C_Prop,
     &                   T,CPVB,NITP,TT,NXX,MXX,
     &                   NinI,TinI,CinI,CT_AVERAGE,N_PW,NFLAG,
     &                   EPSinp,DH0inp)
C---- end modification hokanson 11-Aug-2000: PFPDM09
    
      IMPLICIT NONE

      INTEGER*2 NXX,MXX,Numb,NITP,NinI
      INTEGER*4 N_PW,error
C
C----- Define Workspace----------------------------
C      the HUGE argument should not be needed since the DLL is
C      compiled using the /AH option.
      DOUBLE PRECISION PW[ALLOCATABLE,HUGE](:)
C-------------------------------------------------
C
      DOUBLE PRECISION Chemicals(Numb,7),C_Prop(8),
     &                 Ads_Prop(4),TT(5)
      DOUBLE PRECISION T(400,2),CPVB(Numb,400)
      DOUBLE PRECISION TinI(NinI),CinI(Numb,NinI)
      DOUBLE PRECISION CT_AVERAGE(400)
C---- begin modification hokanson 11-Aug-2000: PFPDM09
      DOUBLE PRECISION EPSinp,DH0inp
C---- end modification hokanson 11-Aug-2000: PFPDM09
      
      INTEGER NCOMPI,MCI,NCI
      PARAMETER(NCOMPI=6,MCI=18,NCI=14)

      INTEGER J,MA,N,ITP

      DOUBLE PRECISION KF(NCOMPI),L,BVF,TCONV,TSTEP,TTOL,TOUT,H0,T0

      DOUBLE PRECISION PARAM(50),QTE,A
      DOUBLE PRECISION TOL
      INTEGER IDO,NSTEPS,NEQ,NFLAG
      INTEGER MC,NC,NCOMP,N1,NIN,MCA,NCA

      DOUBLE PRECISION Y0(750),CBO(NCOMPI),QE(NCOMPI),
     &          ALPHA(NCOMPI,NCOMPI),VALENCE(NCOMPI),
     &          BIS(NCOMPI),BIP(NCOMPI),TIE(4),TINC(4),
     &          VB(NCOMPI),Alpha_Input(NCOMPI),
     &          XWT(NCOMPI),QINIT(NCOMPI)
      DOUBLE PRECISION DG,ST(NCOMPI),DGT,DGS,
     &                 DGP,DP(NCOMPI),DS(NCOMPI),
     &                 EDS(NCOMPI),EDP(NCOMPI)
      DOUBLE PRECISION AZ(MCI,MCI),BR(NCI,NCI),WR(NCI)
      DOUBLE PRECISION AZ1(MCI,MCI),BR1(NCI,NCI),WR1(NCI)
      DOUBLE PRECISION D(NCOMPI),XN(NCOMPI),XNI(NCOMPI),YM(NCOMPI)
      DOUBLE PRECISION CIN(NCOMPI,400),TIN(400)
      DOUBLE PRECISION FMIN(NCOMPI),TP(400),CP(NCOMPI,400),
     &                 CD(NCOMPI,400),CINT(NCOMPI,400)
      DOUBLE PRECISION RAD,EPOR,RHOP,WT,FLRT,DIA,
     & DH0,DSTEP,DTOL,DT0,TORTU,DOUT,EPS
      DOUBLE PRECISION AREA,BEDVOL,EBED,EBCT,TAU
      DOUBLE PRECISION RK1(NCOMPI),RK2(NCOMPI),RK3(NCOMPI),
     &                 RK4(NCOMPI),TOR,PART,TTORTU
      DOUBLE PRECISION CTOTAL, QTOTAL
      DOUBLE PRECISION CINF
      INTEGER NCOL,NM,I
      INTEGER INDEX, MF
      LOGICAL THE_END
      INTEGER LOOP_START,LOOP_END

      integer tester,filenum,check_flag
      character*12, namefile

      COMMON /BLOCKA/ EDS,EDP,BR,D
      COMMON /BLOCKB/ YM,XNI,XN,WR
      COMMON /BLOCKC/ FMIN,TP,CP,CD,CINT
      COMMON /BLOCKD/ CIN,TIN
      COMMON /BLOCKF/ RK1,RK2,RK3,RK4,TCONV
      COMMON /BLOCKG/ TOR,PART,TTORTU
      COMMON /BLOCKH/ QTOTAL,CTOTAL,ALPHA
      COMMON /BLOCKJ/ CBO
      COMMON /CONST/  NCOMP,MC,NC
      COMMON /DGST/   DG,ST
      COMMON /DGTOT/  DGT
      COMMON /MATAZ/  AZ
      COMMON N1,NIN

      common /blocki/tester,itp

C      EXTERNAL FCN,FCNJ,IWKIN
C      DATA INDEX/1/, MF/22/
      DATA NSTEPS/400/

      NFLAG=0
C
C----- Workspace allocation ------------------------------------------
      ALLOCATE (PW(N_PW),STAT=error)
      IF (error.NE.0) GOTO 9999

C
C----- Initialize Check_Flag
C-----   Check_Flag = 0 to Print no Check Files
C-----   Check_Flag = 1 to Print the check files:  INPUT.CHK ALPHA.CHK
C-----                  PFPDM11.CHK REACHINF.CHK 

      check_flag = 1

C
C----- Variables initialization - Values from Fortran = Values from VB
C
      NCOMP = Numb
      DO 801 I = 1,NCOMP
	XWT(I)= Chemicals(I,1)
	CBO(I)= Chemicals(I,2)
	Alpha_Input(I) = Chemicals(I,3)
	KF(I) = Chemicals(I,4)
	DP(I) = Chemicals(I,5)
	VALENCE(I) = Chemicals(I,6) 
	QINIT(I) = Chemicals(I,7)
  801 CONTINUE

C---- Carbon Properties

      EPOR  = C_Prop(1)
      RHOP  = C_Prop(2)
      RAD   = C_Prop(3)
      TORTU = C_Prop(4)
      TOR   = C_Prop(5)
      PART  = C_Prop(6)
      TTORTU= C_Prop(7)
      QTOTAL = C_Prop(8)  
      
C----- Adsorber Data
C     Note : L,DIA are converted from Meters to Centimeters
C            WT from Kilograms to Grams
C            FLRT from m3/s to ml/min

      L    = Ads_Prop(1)*100.D0
      DIA  = Ads_Prop(2)*100.D0
      WT   = Ads_Prop(3)*1000.D0
      FLRT = Ads_Prop(4)*60.D0*1D6

      if (check_flag.ne.0) then
	 open(unit=9, file='input.chk')

	 write(9,*) 'Numb = ',Numb
	 write(9,*) 'NCOMP = ',NCOMP
	 write(9,*)
	 write(9,*)
	 write(9,*) 'Component Properties:'
	 write(9,*)
	 write(9,*)
	 do 2248, i = 1,NCOMP
	    write(9,*) 'Chemical #',i
	    write(9,*)
	    write(9,*) 'XWT(i) (mg/mmol) = ',XWT(i)
	    write(9,*) 'CBO(i) (mg/L)= ',CBO(i)
	    write(9,*) 'Alpha_Input(i) = ',Alpha_Input(i)
	    write(9,*) 'kf(i) (cm/s) = ',KF(i)
	    write(9,*) 'Dp(i) (cm2/s) = ',DP(i)
	    write(9,*) 'VALENCE(i) = ',VALENCE(i)
	    write(9,*)
	    write(9,*)
 2248    continue
	 write(9,*)
	 write(9,*) 'Resin Properties:'
	 write(9,*)
	 write(9,*) 'EPOR (-) = ',EPOR
	 write(9,*) 'RHOP (g/cm3) = ',RHOP
	 write(9,*) 'RAD (cm) = ',RAD
	 write(9,*) 'TORTU (-) = ',TORTU 
	 write(9,*) 'QTOTAL (meq/g) = ',QTOTAL
	 write(9,*)
	 write(9,*)
	 write(9,*)
	 write(9,*) 'Bed Properties:'
	 write(9,*)
	 write(9,*) 'L (cm) = ',L
	 write(9,*) 'DIA (cm) = ',DIA
	 write(9,*) 'WT (g) = ',WT
	 write(9,*) 'FLRT (ml/min) = ',FLRT
	
	 close(unit=9)
      end if

C------ Calculate Two-Dimensional Array of Separation Factors (ALPHA)
C       that will be used in program calculations

      DO 1000, I = 1,Numb
	 DO 1010, J = 1,Numb
	     ALPHA(I,J) = Alpha_Input(I)/Alpha_Input(J)
 1010    CONTINUE
 1000 CONTINUE

      if (check_flag.ne.0) then
	 open(unit=12,file='alpha.chk')

	 do 1020, I = 1,Numb
	    do 1030, J = 1,Numb
	       write(12,*) 'I = ', I, '  J = ', J, 
     &                        '   ALPHA(I,J) = ', ALPHA(I,J)
 1030       continue
	    write(12,*)
 1020    continue

	 close(unit=12)
      end if


C------ Parameters

C---- begin modification hokanson 11-Aug-2000: PFPDM09
C      EPS  = 5.0D-04
C      DH0  = 1.0D-9
      EPS = EPSinp
      DH0 = DH0inp
C---- end modification hokanson 11-Aug-2000: PFPDM09

      NCOL = 3
      DT0  = 0.0
      NM   = 0
C   Number of variable influent conc.
      NIN  = NinI 

      DSTEP =TT(3)
      DTOL  =TT(1) 
      DOUT  =TT(2) 
      
C----- END Initialization------------------------
				     
803   IF (NM .EQ. 0) GO TO 811                                          
      TIE(1) = 40000.0D0
      TINC(1)= 2880.0D0
C       READ(4,*) (TIE(I), I = 1 , NM)                                 
C       READ(4,*) (TINC(I), I = 1 , NM)                                

811   IF (NIN .EQ. 0) GO TO 813                                         
      DO 1 J = 1,NIN                                                    
C       READ(4,*) TIN(J), (CIN(I,J), I = 1,NCOMP)                      
	TIN(J)=TinI(J)  
       DO 2 I=1,NCOMP
	CIN(I,J) = CinI(I,J) 
    2  CONTINUE
    1 CONTINUE
							  
C----- Calculate collocation contants

813   CONTINUE 
    
      NC=NXX
      MC=MXX
      NCA=NC
      MCA=MC
      CALL CONSTNT(NCA,MCA,AZ1,BR1,WR1,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      DO 9 I=1,MC
	DO 9 J=1,MC
	 AZ(I,J)=AZ1(I,J)
 9    CONTINUE
      DO 8 I=1,NC
	WR(I)=WR1(I)
	DO 8 J=1,NC
	  BR(I,J)=BR1(I,J)
8     CONTINUE

      NEQ = (MC*(NC + 1) - 1)*(NCOMP-1)                                     
C                                                                       
C.....calculate fixed bed parameters for LIQUID PHASE                                    
C                                                                       
      DO 212, I = 1, NCOMP                                             
	 CBO(I) = CBO(I)*VALENCE(I)/XWT(I)                                         
	 DO 211 J=1,NIN
	  CIN(I,J)=CIN(I,J)*VALENCE(I)/XWT(I)
211    CONTINUE
212   CONTINUE                                                          
      AREA = 3.141592654D0*DIA*DIA/4.0D0                                
      BEDVOL = L*AREA                                                   
      EBED = 1.0D0 - WT/(BEDVOL*RHOP)                                   
      EBCT = BEDVOL/FLRT                                                
      TAU = BEDVOL*EBED*60.0D0/FLRT                                     


C***  Calculation of Dimensionless Groups for PFPDM with
C***  Dimensionless Groups based on QTOTAL and CTOTAL, where
C***  CTOTAL = Sum of CBO(i) for i = 1 to NCOMP.  Note that
C***  this method has a Langmuir for ION EXCHANGE as the equilibrium
C***  description.
C                                                                       
C.....calculate and print out dimensionless groups                      
C                                                                       
	 DGP = EPOR*(1.0 - EBED)/EBED                                

      CTOTAL = 0.0D0
      DO 30 I = 1,NCOMP                                                 
	 CTOTAL = CTOTAL + CBO(I)
	 EDP(I) = DP(I)*DGP*TAU/(RAD**2)                             
	 ST(I)  = KF(I)*(1.0 - EBED)*TAU/(EBED*RAD)                     
	 BIP(I) = ST(I)/EDP(I) 
   30 CONTINUE                                                          

	 DGS = (RHOP*QTOTAL*(1.0 - EBED)*1000.0)/(EBED*CTOTAL)        
	 DGT = DGS + DGP
	 DG = DGT
							  
      BVF = EBED*(DGT+1.0D0)                                                    

      if (check_flag.ne.0) then
       open(unit=9,file='pfpdm11.chk')
       
       write(9,*) 'Bed Parameters:'
       write(9,*)
       write(9,*) 'Area (cm2) = ',AREA
       write(9,*) 'Volume (cm3) = ', BEDVOL
       write(9,*) 'Bed Porosity (-) = ',EBED
       write(9,*) 'EBCT (min) = ',EBCT
       write(9,*) 'Tau (sec) = ',TAU
       write(9,*)
       write(9,*)
       write(9,*)
       write(9,*) 'Dimensionless Groups based on:'
       write(9,*)
       write(9,*) '        Qtotal (meq/L) = ',QTOTAL
       write(9,*) '            CT (meq/L) = ',CTOTAL
       write(9,*)
       write(9,*)
C---- begin modification hokanson 11-Aug-2000: PFPDM09
       write(9,*)
       write(9,*)
       write(9,*) '                  Error Criteria, EPS = ',EPS
       write(9,*) 'Initial Time Step for Integrator, DH0 = ',DH0
       write(9,*)
       write(9,*)
C---- end modification hokanson 11-Aug-2000: PFPDM09
       write(9,*)
       write(9,*)
       write(9,*)

       do 7000, i = 1,NCOMP
	  write(9,*) 'I = ',I
	  write(9,*)
	  write(9,*) 'DGS = ',DGS
	  write(9,*) 'DGP = ',DGP
	  write(9,*) 'DGT = ',DGT
	  write(9,*)
	  write(9,*) 'EDP(I) = ',EDP(I)
	  write(9,*) 'ST(I) = ',ST(I)
	  write(9,*) 'BIP(I) = ',BIP(I)
	  write(9,*)
	  write(9,*)
	  write(9,*)
 7000  continue

       close(unit=9)
      end if


C                                                                       
C.....call subroutine ORTHOG to combine collocation constants           
C.....and dimensionless groups and to determine total number            
C.....of differential equations being solved for by DGEAR                
C                                                                       
      CALL ORTHOG (N)
C                                                                       
C.....convert independent variables to dimensionless form               
C                                                                       
      TCONV = 60.0/(TAU*(DGT + 1))                                      
      TSTEP = DSTEP*TCONV                                               
      TTOL  = DTOL*TCONV                                                
      TOUT  = DOUT*TCONV                                                
      H0    = DH0*TCONV                                                 
      T0    = DT0*TCONV
						 
      DO 40 I = 1,NM                                                    
	 TIE(I)  = TIE(I)*TCONV                                         
	 TINC(I) = TINC(I)*TCONV                                        
   40 CONTINUE                                                          
C                                                                       
C.....convert influent and experimental data to dimensionless form      
C                                                                       
      DO 60 J = 1,NIN                                                   
	 TIN(J) = TIN(J)*TCONV                                          
	 DO 55 I = 1,NCOMP                                              
	    CIN(I,J) = CIN(I,J)/CTOTAL                                  
   55    CONTINUE                                                       
   60 CONTINUE

							  
C                                                                       
C.....initialize dependent variables 
C.....   This case assumes that there may be more than one presaturant
C.....   on the resin.  The ion specified by the user as the presaturant
C.....   is the last item in the list and is not solved by an ODE.  
C.....   However, if any other ion has presaturant, then the radial
C.....   positions need to be initialized appropriately.
C                                                                       

      DO 66,I=1,N
	 Y0(I) = 0.0D0
   66 CONTINUE

      DO 67  I = 1,(NCOMP-1)                                                   
	 LOOP_START = (I-1)*N1 + 1
	 LOOP_END = LOOP_START + MC*NC - 1
	 DO 68, J = LOOP_START,LOOP_END
	    Y0(J) = QINIT(I)                                                    
   68    CONTINUE
   67 CONTINUE                                                          
   
      if (check_flag.ne.0) then
	 open(unit=12,file='Y0INIT.CHK')
	    do 73, I = 1,N
	       write(12,*) 'I = ',I,'Y0(I) = ',Y0(I)
   73       continue
	 close(unit=12)
      end if
   

C                                                                       
C.....loop for calling GEAR to integrate differential equations         
C                                                                       
C     The two following parameters must be set for the 
C         first call to DGEAR - See Doc.
      INDEX = 1
      MF  = 22

C---- begin modification hokanson 12-Aug-2000: PFPDM10
      open (unit=14, file='PDM11MG.OUT')
      open (unit=15, file='PDM11MEQ.OUT')
      open (unit=16, file='PDM11DIM.OUT')

      write(14,*) 'Column 1: Time in minutes'
      write(14,*) 'Other Columns: Concentrations in mg/L'
      write(14,*)
      write(14,*)

      write(15,*) 'Column 1: Time in Days'
      write(15,*) 'Other Columns: Concentrations in meq/L'
      write(15,*)
      write(15,*)

      write(16,*) 'Column 1: Time in Minutes'
      write(16,*) 
     &  'Other Columns: Dimensionless Concentrations, C/CT'
      write(16,*)
      write(16,*)

C---- end modification hokanson 12-Aug-2000: PFPDM10

      THE_END=.FALSE.
      ITP = 0
      MA  = 1
   70 ITP = ITP + 1

      tester = 0

C
C------------------------- Call GEAR  -------------------------------
C
      IF (ITP.EQ.1) THEN
	IDO=1
      ELSE IF ((ITP.EQ.NSTEPS).OR.THE_END) THEN
	IDO=3
      ELSE
	IDO=2
      ENDIF
       TOL=1.D-4
       PARAM(1)=H0
       PARAM(12)=2
       PARAM(13)=2
       PARAM(19)=0

C       CALL DIVPAG(IDO,N,FCN,FCNJ,A,T0,TOUT,TOL,PARAM,Y0,NFLAG)

      CALL DGEAR (N,T0,H0,Y0,TOUT,EPS,MF,INDEX,NFLAG,PW,N_PW)
C      
      IF (NFLAG.NE.0) GOTO 9999

      IF (NIN.EQ.0) THEN
	 CP(NCOMP,ITP) = 1.0D0
      ELSE
	 CP(NCOMP,ITP) = 0.0D0
	 DO 72,I=1,NCOMP
	    CP(NCOMP,ITP) = CP(NCOMP,ITP) + CINF(I,TOUT)
 72      CONTINUE
      END IF
      CT_AVERAGE(ITP) = CP(NCOMP,ITP) 

      DO 75 I = 1,(NCOMP-1)                                                 
	 CP(I,ITP) = Y0(N1*I)                                           
	 CP(NCOMP,ITP) = CP(NCOMP,ITP) - CP(I,ITP)
   75 CONTINUE                                                          
      TP(ITP) = TOUT

C----- Begin Modification Hokanson 12-Aug-2000: PFPDM10
C--- Write Results for Current Time Step to Output Files
C---    Unit 14='PDM11MG.OUT': mg/L versus BVT
C---    Unit 15='PDM11MEQ.OUT': meq/L versus Days
      
        WRITE(14,8000) TOUT/TCONV,
     &       (CP(I,ITP)*CTOTAL*XWT(I)/VALENCE(I),I=1,NCOMP)
        WRITE(15,8000) TOUT/TCONV/1440.0D0,
     &       (CP(I,ITP)*CTOTAL,I=1,NCOMP)
        WRITE(16,8000) TOUT/TCONV,
     &       (CP(I,ITP),I=1,NCOMP)
 8000   FORMAT(1X,10(E20.12))                
C----- End Modification Hokanson 12-Aug-2000: PFPDM10                                          

      DOUT = TOUT/TCONV                                                 
      IF ( ITP .LT. NSTEPS ) THEN                                       
	 IF ( TOUT .LT. TTOL ) THEN                                     
	    IF ( NM .NE. 0 .AND. TOUT .GE. TIE(MA) ) THEN               
	       TSTEP = TINC(MA)                                         
	       IF ( MA .EQ. NM ) THEN                                   
		  NM = 0                                                
	       ELSE                                                     
		  MA = MA + 1                                           
	       ENDIF                                                    
	    ENDIF                                                       
	    TOUT = TOUT + TSTEP                                         
	    IF ( TOUT .GT. TTOL ) THEN
	       TOUT = TTOL                           
	       THE_END=.TRUE.
	    ENDIF
	    GO TO 70                                                    
	 ENDIF                                                          
      ELSE                                                              
	 IF ( TOUT .NE. TTOL ) THEN                                     
	    GO TO 9999                                                    
	 ENDIF                                                          
      ENDIF
C 
C----- Transfer data ---------------------------------                                                            
C
 90   DO 82 J = 1 , ITP                                                 
	T(J,1)=TP(J)/TCONV      
	T(J,2)=TP(J)*BVF
	DO 821 I=1,NCOMP
	   CPVB(I,J)=CP(I,J)
 821    CONTINUE
 82   CONTINUE                                                          
      NITP=ITP

C---- begin modification hokanson 12-Aug-2000: PFPDM10
      close (16)
      close (15)
      close (14)
C---- end modification hokanson 12-Aug-2000: PFPDM10
      
9999  IF (error.NE.0) then 
       NFLAG = 1603 
      ELSE
       DEALLOCATE(PW,STAT=error)
       IF (error.NE.0) THEN 
	 NFLAG = 1603
       ENDIF
      ENDIF
      RETURN
      END                                                               
C                                                                       
C                                                                       
C                      -----------------------                          
C                      I END OF MAIN PROGRAM I                          
C                      -----------------------                          
C                                                                       
C                                                                       
C                                                                       
C
C                                                                       
C   **************************************************************      
C   * This subroutine combines the collocation constants and the *      
C   * dimensionless groups calculated in the main program to     *      
C   * save computation time.                                     *      
C   **************************************************************      
C

      SUBROUTINE ORTHOG(N)
      IMPLICIT NONE

      INTEGER NCOMPI,MCI,NCI
      PARAMETER(NCOMPI=6,MCI=18,NCI=14)

      INTEGER N,NCOMP,MC,NC
      INTEGER MD,MND,ND,NIN,N1,I,J,K

      DOUBLE PRECISION DG,ST(NCOMPI),EDS(NCOMPI),EDP(NCOMPI),
     & DGI,DGT,D(NCOMPI),STD(NCOMPI),BEDS(NCOMPI,NCI,NCI),
     & BEDP(NCOMPI,NCI,NCI)
      DOUBLE PRECISION BR(NCI,NCI)

      DOUBLE PRECISION EDD(NCOMPI)

      COMMON /BLOCKA/ EDS,EDP,BR,D
      COMMON /BLOCKE/ STD,BEDS,BEDP,DGI,MND,ND,MD
      COMMON /CONST/  NCOMP,MC,NC
      COMMON /DGST/   DG,ST
      COMMON /DGTOT/  DGT
      COMMON N1,NIN



      ND  = NC - 1                                                      
      MD  = MC - 1                                                      
      MND = MC*ND                                                       
      N1  = MND + MC + MD                                               
      N   = N1*(NCOMP-1)                                                    
      DO 50 I = 1,(NCOMP-1)
	 DGT    = 1.0 + DGT                                             
	 DGI = 1.0/DG                                             
	 STD(I) = ST(I)*DGT                                             
	 EDD(I) = DGT/DG                                           

	 DO 40 J = 1,ND                                                 
	    DO 30 K = 1,NC                                              
	       BEDP(I,J,K) = EDP(I)*EDD(I)*BR(J,K)         
   30       CONTINUE                                                    
   40    CONTINUE                                                       
   50 CONTINUE 
							 
      RETURN                                                            
      END                                                               
C                                                                       
C                                                                       
C                   ----------------------------                        
C                   | END OF SUBROUTINE ORTHOG |                        
C                   ----------------------------                        
C                                                                       
C                                                                       
C--------------------------------------------------------------------
C                       Subroutine DIFFUN
C--------------------------------------------------------------------
C
C  - The system to be solved is Ay'=f(x,y)
C  - Here, A is the identity matrix
C  - FCN is called by DIVPAG  to return the value of f(x,y) for
C    given values of x and y.

      SUBROUTINE DIFFUN(N,T,Y0,YDOT)
      IMPLICIT NONE
      INTEGER NCOMPI,MCI,NCI
      PARAMETER(NCOMPI=6,MCI=18,NCI=14)

      INTEGER I,II,III,IIII,J,JJ,K,KK,N,M,MC,MD,MND,
     &       NC,NCOMP,ND,NIN,N1
      DOUBLE PRECISION Y0(N),YDOT(N),WW(MCI),AAU(MCI),BB(NCI,MCI),
     &   Z(NCOMPI),Q0(NCOMPI),CBS(NCOMPI,MCI),CPORE(750),CINF,
     &   CINFL
      DOUBLE PRECISION AZ(MCI,MCI),BEDP(NCOMPI,NCI,NCI),
     &  BEDS(NCOMPI,NCI,NCI),
     &  DGI,DGT
     & ,QTE,STD(NCOMPI),
     & T,WR(NCI),XN(NCOMPI),XNI(NCOMPI),YM(NCOMPI),YT0
      DOUBLE PRECISION RK1(NCOMPI),RK2(NCOMPI),RK3(NCOMPI),
     &                 RK4(NCOMPI),TCONV,XKTIME(NCOMPI),
     &                 TOR,PART,TORTU,FAC(NCOMPI),RT,TTORTU
      DOUBLE PRECISION EDS(NCOMPI),EDP(NCOMPI),BR(NCI,NCI),
     &                 D(NCOMPI)
      DOUBLE PRECISION CTOTAL,CBO(NCOMPI),QTOTAL,ALPHA(NCOMPI,NCOMPI),
     &                 CTAVG,QPRESAT

      double precision test_dummy, test_dummy2
      integer tester,itp,filenum

      COMMON /MATAZ/ AZ
      COMMON /DGTOT/ DGT
      COMMON /CONST/ NCOMP,MC,NC
      COMMON /BLOCKA/ EDS,EDP,BR,D
      COMMON /BLOCKB/YM,XNI,XN,WR
      COMMON /BLOCKE/STD,BEDS,BEDP,DGI,MND,ND,MD
      COMMON /BLOCKF/RK1,RK2,RK3,RK4,TCONV
      COMMON /BLOCKG/TOR,PART,TTORTU
      COMMON /BLOCKH/QTOTAL,CTOTAL,ALPHA
      COMMON /BLOCKJ/CBO
      COMMON N1,NIN
      
      common /blocki/tester,itp
   

C                                                                       
C.....determine liquid phase concentrations at each radial and          
C.....axial position within adsorbent particle using Ideal              
C.....Adsorbed Solution Theory                                          
C
C                                                                            
C
      DO 2 I=1,NCOMP
C        XKTIME(I) = XK(I)
C        XKTIME(I)= XK(I)*(RK1(I)+RK2(I)*(T/TCONV)
C     &      + RK3(I)*DEXP(RK4(I)*(T/TCONV)))
C        IF (XKTIME(I) .LE. (XK(I)/1.0D+3)) THEN
C           XKTIME(I) = XK(I)/1.0D+3
C        ENDIF
   2  CONTINUE
      DO 3 I = 1,NCOMP
	 FAC(I) = 1.0D0
C      IF (TOR .LT. 1.0D0) THEN
C       IF ((T/TCONV) .GT. TTORTU) THEN
C            TORTU = TOR + PART*(T/TCONV)
C       ELSE
C            TORTU = 1.0D0
C       ENDIF
C       RT = TORTU/1.0D0
C      ELSE
C       TORTU = TOR
C       RT = 1.0D0
C      ENDIF
C      FAC(I) = ((1.0D0/RT) - D(I))/(1.0D0 - D(I))
 3    CONTINUE                                                                        

       

      II = 0                                                            
      JJ = 0                                                            
      DO 15 K = 1,MC                                                    
	 DO 8 M = 1,NC                                                  
	    QPRESAT = 0.0D0
	    DO 5 I = 1,(NCOMP-1)                                            
	       II = II + 1                                              
	       Z(I) = Y0(II)
	       QPRESAT = QPRESAT + Z(I)
	       II = II + N1 - 1                                         
    5       CONTINUE

	    QPRESAT = 1.0D0 - QPRESAT  
	    DO 6 I = 1,(NCOMP - 1)                                           
	       JJ = JJ + 1 
		   
		   QTE = ALPHA(I,NCOMP) * QPRESAT

		   IF (NIN.EQ.0) THEN
		      CTAVG = 1.0D0
		   ELSE
		      CTAVG = 0.0D0
		   END IF

		   DO 111, J = 1,(NCOMP - 1)              
		      QTE = QTE + ALPHA(I,J) * Z(J)
		      IF (NIN.NE.0) THEN
			 CTAVG = CTAVG + CINF(J,T)
		      END IF
 111               CONTINUE
C----- begin modification hokanson 12-Jan-2001: PFPDM11
                   IF (NIN.NE.0) CTAVG = CTAVG + CINF(NCOMP,T)
C----- end modification hokanson 12-Jan-2001: PFPDM11

	       IF ( QTE .LE. 0.0D0 .OR. Z(I) .LE. 0.0D0 ) THEN           
		  CPORE(JJ) = 0.0D0                                     
	       ELSE                                                     
		  CPORE(JJ) = CTAVG * Z(I) / QTE
	       ENDIF                                                 
						   
	       JJ = JJ + N1 - 1                                         
    6       CONTINUE                                                    
	    IF ( M .LT. NC - 1 ) THEN                                   
	       II = (K - 1)*ND + M                                      
	       JJ = (K - 1)*ND + M                                      
	    ELSE                                                        
	       II = (K - 1) + MND                                       
	       JJ = (K - 1) + MND                                       
	    ENDIF                                                       
    8    CONTINUE                                                       
	 II = ND*K                                                      
	 JJ = ND*K                                                      
   15 CONTINUE

      tester = tester + 1
							  
      DO 60 I = 1,(NCOMP-1)                                                 
	 II = (I-1)*N1                                                  
	 III = II + MND                                                 
	 IIII = III + MD                                                
	 IF ( NIN .EQ. 0 ) THEN                                         
	    CINFL = CBO(I)/CTOTAL                                               
	 ELSE                                                           
	    CINFL = CINF(I,T)                                           
	 ENDIF                                                          
	 DO 20 K = 2,MC
	 IF ( CPORE(III + K) .LE. 0.0D0 ) THEN                          
	    CBS(I,K) = STD(I)*Y0(IIII + K)                              
	 ELSE                                                           
	    CBS(I,K) = STD(I)*(Y0(IIII + K) - CPORE(III + K))           
	 ENDIF                                                          
   20    CONTINUE                                                       
	 DO 40 K = 1,MC                                                 
	    WW(K) = 0.0D0                                               
	    AAU(K) = 0.0D0                                              
	    KK = II + (K-1)*ND                                          


	    DO 30 J = 1,ND                                              
	      BB(J,K) = 0.0D0                                           
	       DO 25 M = 1,ND                                           
C----------------------------------------------------------------
C   The factor FAC(I) to account for the variation of tortuosity 
C   is used twice in the three following lines
C----------------------------------------------------------------
		  BB(J,K) = BB(J,K) + BEDP(I,J,M)*FAC(I)*CPORE(KK + M)            
   25          CONTINUE   
	       BB(J,K) = BB(J,K) + BEDP(I,J,NC)*FAC(I)*CPORE(III + K)             
   30       CONTINUE                                                    
	    DO 35 J = 1,ND                                              
	       JJ = KK + J                                              
C                                                                       
C.....Intraparticle Phase Mass Balance (excluding boundary)             
C                                                                       
	       YDOT(JJ) = BB(J,K)                                       
C                                                                       
	       WW(K) = WW(K) + WR(J)*YDOT(JJ)                           
   35       CONTINUE                                                    
   40    CONTINUE                                                       
C                                                                       
C.....Liquid-Solid Boundary Layer Mass Balance at column entrance       
C        
	 YDOT(III+1) = (STD(I)*DGI*(CINFL - CPORE(III + 1))          
     +                 - WW(1)) / WR(NC)                                
C                                                                       
	 DO 55 K = 2,MC                                                 
C                                                                       
C.....Liquid-Solid Boundary Layer Mass Balance within column            
C                                                                       
	    YDOT(III+K) = (CBS(I,K)*DGI - WW(K)) / WR(NC)            
C                                                                       
	    DO 50 M = 2,MC                                              
	       AAU(K) = AAU(K) + AZ(K,M)*Y0(IIII+M)                     
   50       CONTINUE                                                    
C                                                                       
C.....Liquid Phase Mass Balance                                         
C                                                                       
	    YDOT(IIII+K) = -DGT*(AZ(K,1)*CINFL + AAU(K))                
     +                     - 3.0D0*CBS(I,K)                             
C                                                                       
   55    CONTINUE                                                       
   60 CONTINUE       
      RETURN                                                            
      END                                                               
C                                                                       
C                   ----------------------------                        
C                   I END OF SUBROUTINE DIFFUN I                        
C                   ----------------------------                        
C                                                                       
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
      DOUBLE PRECISION FUNCTION CINF(I,T)
      IMPLICIT NONE                              

      INTEGER NCOMPI
      PARAMETER (NCOMPI=6)
      INTEGER I, N1, NIN ,J
      DOUBLE PRECISION CIN(NCOMPI,400),TIN(400),T,CBO(NCOMPI),
     &                 CTOTAL,QTOTAL,ALPHA(NCOMPI,NCOMPI)

      COMMON/BLOCKD/CIN,TIN
      COMMON /BLOCKH/QTOTAL,CTOTAL,ALPHA
      COMMON /BLOCKJ/CBO
      COMMON N1,NIN                                     

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



      SUBROUTINE FCNJ

      END
C--------------------------------------------------------------------
C    Subroutine CONSTANT - Provide Collocation Constants
C--------------------------------------------------------------------
      SUBROUTINE CONSTNT(N1,N2,AZ1,BR1,WR1,nflag)
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      PARAMETER (MCI=18,NCI=14)
      DIMENSION R(18), Z(18), QI(18,18), RR(18)
      DIMENSION AZ1(MCI,MCI),WR1(NCI),BR1(NCI,NCI)
      INTEGER NFLAG
      COMMON/QMTRXS/ Q(18,18), C(18,18), D(18,18), F(18)                
      COMMON/OCCOEF/ AR(18,18), BR(18,18), AZ(18,18), BZ(18,18)         
     +, WZ(18), WR(18)

C       For Spherical coordinates: NGEOR = 3
C                                  NOR =0
C                                  N1R = 1
C                                  ALFAR = 1.D0
C                                  BETAR = 0.5D0
C       For Cylindrical coordinates: NGEOR = 2
C                                    NOR =0
C                                    N1R = 1
C                                    ALFAR = 1.D0
C                                    BETAR = 0.0D0
      DATA NGEOR/3/
     +    ,  N0R/0/,  N1R/1/,  ALFAR/1.0D0/,  BETAR/ 0.5D0/            
     +    ,  N0Z/1/,  N1Z/1/,  ALFAZ/0.0D0/,  BETAZ/ 0.0D0/            
      NFLAG=0
      NR=N1
      NZ=N2                                                  
									
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
9997  return
      END

C **********************************************************************
      SUBROUTINE DROOT(N,N0,N1,AL,BE,ROOT)                              
C **********************************************************************
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)                               
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
      SUBROUTINE PEDERV ( N,T,Y,PD,N0 )                     
C                                                                       
C       ******************************************************          
C       * This subroutine is a dummy subprogram used by GEAR *          
C       ******************************************************          
C                                                                       
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)                               
      COMMON MC,NC,NCOMP,N1,DGT,NIN                                     
      RETURN                                                            
      END                                                               

