
	INCLUDE 'blah.for'

C.......READ IN DATA
        CALL READIN

C.......PERFORM CALCULATIONS
        CALL CALC

C.......OUTPUT RESULTS
        CALL OUTDAT        
	  CALL VBOUTPUT

      STOP
      END


CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE READIN
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
        INCLUDE 'blah.for'
        INTEGER IC

c        OPEN (UNIT=10, FILE='H:\favor\input1.dat')
        OPEN (UNIT=10, FILE='input1.dat')

C *************** PLANT INFLUENT
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 


C Plant Flow Rate, L/day
        READ (10,20) Q
C Solids Influent Concentration, mg/L
        READ (10,20) X0

C *************** PHYSICO-CHEMICAL PROPERTIES
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Barametric Pressure (kPa)      
        READ (10,20) PB
C Temperature (C)
        READ (10,20) T
C Wind Velocity 10 meters about plant, m/s
        READ (10,20) WNDVRI
C Name of Contaminant
        READ (10,30) NAME
C Contaminant Influent Concentration, ug/L
        READ (10,20) CO1
C Biodegradation Rate Constant (L/(mg*day)
        READ (10,20) KB
C Log Octanol Water Coefficient For Contaminant
        READ (10,20) LOGKOW
C Henry's Constant
        READ (10,20) H
C Molecular Weight, g/mol
        READ (10,20) MW
C Diffusivity of Contaminant IN H20, cm2/sec
        READ (10,20) VOCDIF
C Gas Phase Contaminant Diffusivity (cm2/sec)
        READ (10,20) VOCDFG 
C Oxygen Saturation Concentration at Effective Depth (mg/L)
        READ (10,20) CSAT
C Henry's Constant For Oxygen
        READ (10,20) HOXY
C Diffusivity of Oxygen, cm2/sec
        READ (10,20) OXYDIF
C Density of Water (kg/m3)
        READ (10,20) H2ODEN
C Viscosity of Water (kg/(m*s))
        READ (10,20) H2OVIS
C Vapor Pressure (kPa)      
        READ (10,20) PV
C Process Water Correction (ALPHA)
        READ (10,20)  ALPHA
C Denisty of Air(kg/m3)      
        READ (10,20) AIRDEN
C Viscosity of Air (kg/(m*s))
        READ (10,20) AIRVIS 

C *************** PRIMARY CLARIFIERS
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Number of Primary Clarifiers being Modeled
        READ (10,27) NPC
C Covered Primary Clarifier Option (0=off, -1=on)
        READ (10,27) CBPC
C Primary Clarifier Ventilation Air Flow Rate for each Clarifier 
        READ (10,20) QVPC
C Primary Clarifier Basin Depth for each, m
        READ (10,20) PCD
C Primary Clarifier Volume for each, L
        READ (10,20) PV1
C Primary Wastage Flow Rate from each Clarifier, L/day
        READ (10,20) QW1
C Percent Solids Removal in Primary Clarifier 
 	  READ (10,20) E
C Primary Sorption Mechanism
        READ (10,27) PSM
C Primary Volitalization Mechanism (1=McKay & Yeun, 2=KLA)
        READ (10,27) PVM

C *************** INFLUENT WEIR
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Option for Influent Weir Drop (0=off -1=on)
        READ (10,27) W1
C Influent Weir Mechanism (1=NAPPE 2=POOL)
        READ (10,27) WM1
C Width of Influent Weir Channel (m)
        READ (10,20) WW1
C Distance Between Water Levels Above and Below Influent Weir (M)
        READ (10,20) Z1
C Gas Flow Rate Leaving the Tailwater per Unit Influent Weir Length (m3/(m*h))
        READ (10,20) QG1

C *************** PRIMARY CLARIFIER WEIR
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Option for Secondary Influent Weir Drop (0=off -1=on)
        READ (10,27) W2
C Primary Effluent Weir Mechanism (1=NAPPE 2=POOL) 
        READ (10,27) WM2
C Width of the Weir Channel (m)
        READ (10,20) WW2
C Distance Between Water Levels Above and Below Weirs (m)
        READ (10,20) Z2
C Gas Flow Rate Leaving the Tailwater per Unit Weir Length (m3/(m*h))
        READ (10,20) QG2

C *************** SECONDARY CLARIFIER WEIR
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Option For Effluent Drop (0=off -1=on)
        READ (10,27) W3
C Secondary Effluent Weir Mechanism (1=NAPPE 2=POOL)
        READ (10,27) WM3
C Width of the Weir Channel (m)
        READ (10,20) WW3
C Distance Between Water Levels Above and Below Weir (m)
        READ (10,20) Z3
C Gas Flow Rate Leaving Tailwater per Unit Weir Length (m3/(m*h))
        READ (10,20) QG3

C *************** AERATED GRIT CHAMBER
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Option for Modeling of Aerated Grit Chamber (0=off -1=on) 
        READ (10,27) GC
C Number of Aerated Grit Chambers being Modeled
        READ (10,27) NAGC
C Covered Aerated Grit Chamber (0=off, -1=on)
        READ (10,27) CBAGC
C Aerated Grit Chamber Total Ventilation Air Flow Rate for each Chamber,
        READ (10,20) QVAGC
C Aerated Grit Chamber Depth for each (m)
        READ (10,20) AGCD
C Aerated Grit Chamber Volume for each (L)
        READ (10,20) AGCV
C Gas Flow Rate For each Aerated Grit Chamber (L/min)
        READ (10,20) QGGC 
C SOTR For Bubble Aeration in Aerated Grit Chamber (kg/hr),
        READ (10,20) AGBSOT

C *************** AERATION BASINS
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Number of Aeration Basins being Modeled
        READ (10,27) NAB
C Covered Aeration Basin Option (0=off, -1=on)
        READ (10,27) CBAB
C Ventilation Air Flow Rate for each Aeration Basin, 
        READ (10,20) QV
C Aeration Basin Depth (m),
        READ (10,20) ABD
C Sludge Wastage From Each Secondary Clarifer, SQW (L/day)             
        READ (10,20) SQW
C Recycle Flow Rate from each Secondary Clarifier, L/day
        READ (10,20) QR
C Secondary Aeration Mechanism (1=SURFACE 3=DIFFUSED BUBBLE)
	  READ (10,27) SAM 

	  IF ( SAM .EQ.1) THEN
C SOTR For Surface Aeration (kg/hr)
            READ (10,20) SUFSOR
	  ELSE
C SOTR For Bubble Aeration in Aeration Basin (kg/hr),
            READ (10,20) ABBSOT
	  ENDIF

C *************** SECONDARY CLARIFIER
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C Number of Secondary Clarifier being Modeled
        READ (10,27) NSC
C Covered Secondary Clarifier Option (0=off, -1=on)
        READ (10,27) CBSC
C Secondary Clarifier Ventilation Air Flow Rate for each Clarifier,
        READ (10,20) QVSC
C Secondary Clarifier Basin Depth for each (m),  
        READ (10,20) SCBD
C Secondary Clarifier Basin Volume for each (L)
        READ (10,20) SCBV
C Solids Concentration in Secondary Effluent (mg/L)
	   READ (10,20) XSC


        CLOSE(UNIT=10)
       
C *************** AERATION BASINS CSTRs

c        OPEN (UNIT=14, FILE='H:\favor\input2.dat')
        OPEN (UNIT=14, FILE='input2.dat')
C Step Feed Modeling (0=off 1=on)	
        READ (14,27) SF 
C Number of Tanks Being Modeled       
        READ (14,27) NTK


C Biomass Concentration For A Particular Tank, XBM (mg/L)             
C Aeration Tank Volume For A Particular Tank, ATV (L)             
C Gas Flow Rate For a Particular Tank, QG (L/min)             
C Fraction of Plant Influent Directly Entering, CSTR FFRACT
        DO 50 IC=1, NTK
            READ (14,20) XBM(IC)
            READ (14,20) ATV(IC)
            READ (14,20) QG(IC)
            READ (14,20) FFRACT(IC)
   50   CONTINUE

        CLOSE(UNIT=14)

20    FORMAT (/ E12.5) 
27  	FORMAT (/ I5)       
30    FORMAT (/ A20)
        RETURN
        END




CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE CALC
C                                                                   C
C       DUMMY SUBROUTINE FOR CALLING ALL SUBROUTINES THAT           C
C       INVOLVE CALCULATIONS
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
        INCLUDE 'blah.for'
c        COMMON A(10,10),B(10),IPVT(10)

C.......READ IN DATA
        CALL PARAMS

C.......INFLUENT WEIR MASS BALANCE
        CALL INWEIR

C.......AERATED GRIT CHAMBER MASS BALANCE
        CALL AERGRT

C.......PRIMARY CLARIFIER MASS BALANCE
        CALL PRIMAR

C.......SECONDARY TREATMENT: MASS BALANCE FOR CONVENTIONAL
C.......ACTIVATED SLUDGE WITH STEP FEED OPTION INCLUDED
        CALL SECCON

        CALL REMOV

        RETURN
        END


CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE PARAMS
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
        INCLUDE 'blah.for'

C.....CONVERT X0 TO VSS FROM TSS, ASSUMING THAT VSS/TSS = 0.8
      X0=X0*0.8
	XSC=XSC*0.8

C.....CALCULATION OF KP1 
      IF (PSM.EQ.1) THEN
         LOGKP1=5.8D-1*(LOGKOW)+1.14D0
         KP1=10.D0**LOGKP1
	   KP1=KP1/1000.D0
      ELSE
         LOGKP1=0.67D0*(LOGKOW)-2.61D0
         KP1=10.D0**LOGKP1
C........CONVERT KP1 FROM TSS BASIS TO VSS BASIS, ASSUMING
C........THAT VSS/TSS = 0.8
         KP1=KP1/0.8	   
      END IF
      KP1=KP1/1000.D0
C.....CONVERT KP1 FROM L/GM TO L/MG

C
C.....STORE THIS VALUE FOR OUTPUT.
	KP1_OUT = KP1
C.....END OF STORE THIS VALUE FOR OUTPUT.
C

C.....CALCULATION OF SPECIFIC WEIGHT OF WATER (kPa/m)
      H2OSW=H2ODEN*9.81D0/1.D3
           
C.....ASSIGNMENT OF AIR MOLECULAR WEIGHT (g/gmol)
      AIRMW=2.88D+1
           
C.....CALCULATION OF SCHMIDT NUMBERS
         SCGPRI=AIRVIS*1.D+04/(VOCDFG*AIRDEN)
C........CONVERT DIFFUSIVITY FROM CM2/S TO M2/S
         SCLPRI=H2OVIS*1.D+04/(VOCDIF*H2ODEN)
C........CONVERT DIFFUSIVITY FROM CM2/S TO M2/S

C.....CALCULATE TOTALED PARAMETERS
         TATV = 0.D0
         TQG = 0.D0
         DO 100 IC=1,NTK
            TATV = TATV + ATV(IC)
            TQG = TQG + QG(IC)
  100    CONTINUE

      RETURN
      END

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE INWEIR
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      INCLUDE 'blah.for'

      IF (W1.EQ.-1) THEN
         RV1=WEIRV(WM1,Z1,Q,WW1,QG1)
      ELSEIF (W1.EQ.0) THEN 
         RV1=1.D0
      END IF

      FE1=1.D0-(1.D0/RV1)
      CA=CO1*(1.D0-FE1)
      RVOLW1=Q*CO1*FE1

C
C.....STORE THIS VALUE FOR OUTPUT.
	XVALS_OUT(1) = X0
C.....END OF STORE THIS VALUE FOR OUTPUT.
C
      RETURN
      END

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE AERGRT
C                                                                   C
C       AERATED GRIT CHAMBER                                        C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      INCLUDE 'blah.for'

      IF (GC.EQ.-1) THEN

         QAGC=Q/DFLOAT(NAGC)

         XKLAVL=KLAVL(CBAGC,QVAGC,AGCD,AGCV)

         XBUB=BUBPAR(AGBSOT,QGGC,AGCV,AGCD)

C........COVERED BASIN CALCS
         IF(CBAGC.EQ.-1) THEN
            D3N = XKLAVL*AGCV
            D5N = H*QGGC*60.D0*24.D0*XBUB
C...........CONVERT QGGC FROM L/MIN TO L/DAY
            AG=((D3N+D5N)/((QVAGC*24.D0*60.D0)+(D3N/H)))/H
C...........CONVERT QVAGC FROM L/MIN TO L/DAY
         ELSE
            AG=0.D0
         ENDIF

C........MASS BALANCE
         CB=(QAGC*CA*(1.D0+(KP1*X0)))/
     1     ( (QAGC*(1.D0+(KP1*X0)))
     2     + (XBUB*QGGC*H*60.D0*24.D0)
C...........CONVERT QGGC FROM L/MIN TO L/DAY
     3     + (XKLAVL*(1.D0-AG)*AGCV) )

         RSTRPG=XBUB*QGGC*H*60.D0*24.D0*CB
     1          *DFLOAT(NAGC)
         RVOLAG=XKLAVL*(1.D0-AG)*AGCV*CB
     1          *DFLOAT(NAGC)

      ELSEIF (GC.NE.1) THEN 
         CB=CA
      END IF
      

C
C.....STORE THIS VALUE FOR OUTPUT.
	XVALS_OUT(2) = X0
C.....END OF STORE THIS VALUE FOR OUTPUT.
C
      RETURN
      END

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE PRIMAR
C                                                                   C
C       PRIMARY CLARIFIER                                           C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      INCLUDE 'blah.for'

C.....1) PRIMARY CLARIFIER

         QPC=Q/DFLOAT(NPC)

         QE1=QPC-QW1

         XKLAVL=KLAVL(CBPC,QVPC,PCD,PV1)

         XW1=E*QPC*X0/QW1
	   XPC=(1.D0-E)*QPC*X0/QE1

C........COVERED BASIN CALCS
         IF(CBPC.EQ.-1) THEN
C...........INCLUDE GAS FROM PRIMARY EFFLUENT WEIR DROP
            IF (W2.EQ.-1) THEN
               RV2=WEIRV(WM2,Z2,QE1,WW2,QG2)
            ELSEIF (W2.EQ.0) THEN 
               RV2=1.D0
            END IF
            FE2=1.D0-(1.D0/RV2)
            D3N = XKLAVL*PV1
            AG=((D3N+(QE1*FE2))/
     1         ((QVPC*24.D0*60.D0)+(D3N/H)))/H
C...........CONVERT QVPC FROM L/MIN TO L/DAY
         ELSE
            AG=0.D0
         ENDIF

C........MASS BALANCE
         CC=(QPC*CB*(1.D0+(KP1*X0)))/
     1      (QE1*(1.D0+(KP1*XPC))+(QW1*(1.D0+(KP1*XW1)))+
     2      (XKLAVL*(1.D0-AG)*PV1))

         RSORPI=(QW1*KP1*XW1)*CC*DFLOAT(NPC)
         RSORPW=QW1*CC*DFLOAT(NPC)
         RVOLPRI=(XKLAVL*(1.D0-AG)*PV1)*CC*DFLOAT(NPC)

C.....2) PRIMARY EFFLUENT WEIR DROP
         IF (W2.EQ.-1) THEN
            RV2=WEIRV(WM2,Z2,QE1,WW2,QG2)
         ELSEIF (W2.EQ.0) THEN 
            RV2=1.D0
         END IF

         FE2=1.D0-(1.D0/RV2)
         CD=CC*(1.D0-FE2)
         RVOLW2=QE1*CC*FE2*DFLOAT(NPC)
  
C
C.....STORE THIS VALUE FOR OUTPUT.
	XVALS_OUT(3) = XPC
	XVALS_OUT(4) = XPC
C.....END OF STORE THIS VALUE FOR OUTPUT.
C
      RETURN
      END

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE SECCON
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
        INCLUDE 'blah.for'
        INTEGER IC,JC,IPVT,SAMX
        DIMENSION A(10,10),B(10),IPVT(10)
        DIMENSION AGA(10),XKLASU(10),XBUB(10)
        DIMENSION RBIOSC(10),RSUFSC(10),RBBLSC(10),RVLSCA(10)

C......ADD 1 TO NUMBER OF CSTRS FOR LOOPS INVOLVING
C......SECONDARY CLARIFIER       
       NTKP1=NTK+1

C......IF STEP FEED IS NOT UTILIZED (SF=0), 
C......SET FFRACT(IC) TO APPROPRIATE VALUES
       IF ((SF.EQ.0).OR.(NTK.EQ.1)) THEN
          FFRACT(1)=1.D0
          IF(NTK.GT.1) THEN
             DO 100 IC=2,NTK
100          FFRACT(IC)=0.D0
          END IF
       END IF
   

C.....INITIALIZE MATRIX AND VECTORS
      DO 90 IC=1,NTKP1
         B(IC)=0.D0
         CE2(IC)=0.D0
         DO 90 JC=1,NTKP1
            A(IC,JC)=0.D0
   90 CONTINUE
      DO 95 IC=1,NTK
         AGA(IC)=0.D0
         XKLASU(IC)=0.D0
         XBUB(IC)=0.D0
         RBIOSC(IC)=0.D0
         RSUFSC(IC)=0.D0
         RBBLSC(IC)=0.D0
         RVLSCA(IC)=0.D0
   95 CONTINUE

C.....BEGIN COMPUTING MATRIX AND VECTOR COEFFICIENTS

C.....INITIALIZE AND SET PARAMETERS THAT DON'T DEPEND
C.....ON INDIVIDUAL CSTRS
      SAMX=SAM
      QAERDIV=(QE1*DFLOAT(NPC))/DFLOAT(NAB)
      QRDIV=QR*DFLOAT(NSC)/DFLOAT(NAB)
      QTDIV = QAERDIV + QRDIV
      QN = 0.D0   
      XBMM1 = 0.D0   
      QSC=(QTDIV*DFLOAT(NAB))/DFLOAT(NSC)
      XR=(XBM(NTK)*QSC-(QSC-QR-SQW)*XSC)/(QR+SQW)

C.....VOLATILIZATION FOR AERATION BASIN 
      XKLAV1=KLAVL(CBAB,QV,ABD,TATV)

C.....BEGIN CSTR LOOP FOR AERATION BASIN
      DO 110 IC=1,NTK

C........SET UP COEFFICIENTS FOR AERATION CSTRS

C........FLOWRATES, BIOMASS CONCNS, TANK VOLUMES
         QNF=FFRACT(IC)*QTDIV
         QNM1=QN
         QN=QN+QNF
         QGN=QG(IC)
         IF (IC.GT.1) XBMM1=XBM(IC-1)
         XBMT=(XR*QRDIV+XPC*QAERDIV)/QTDIV
         XBMN=XBM(IC)
         XBMW=XR
         ATVN=ATV(IC)

C........MASS TRANSFER FOR SURFACE OR BUBBLE AERATION
C........SURFACE AERATION
         IF (SAMX.EQ.1) THEN
            XBUB(IC)=0.D0
            XKLASU(IC)=KLASUF(ATVN,ABD)
C........DIFFUSED BUBBLE AERATION}
         ELSE IF (SAMX.EQ.3) THEN
            XBUB(IC)=BUBPAR(ABBSOT,QGN,ATVN,ABD)
            XKLASU(IC)=0.D0
         END IF

C........COMPONENTS OF COEFFICIENTS
         DNM1 = QNM1*(1.D0+(KP1*XBMM1))
         DN = QNF*(1.D0+(KP1*XBMT))
         D1N = QN*(1.D0+(KP1*XBMN))
         D2N = KB*XBMN*ATVN
         D3N = XKLAV1*ATVN
         D4N = XKLASU(IC)*ATVN
         D5N = H*QGN*60.D0*24.D0*XBUB(IC)
C........CONVERT QGN FROM L/MIN TO L/DAY
         IF(CBAB.EQ.-1) THEN
            AGA(IC)=((D3N+D4N+D5N)/
     1              ((QV*24.D0*60.D0)+(D3N/H)+(D4N/H)))/H
C...........CONVERT QV FROM L/MIN TO L/DAY
         ELSE
            AGA(IC)=0.D0
         ENDIF
         AT1=QAERDIV*CD*(1.D0+(KP1*XPC))/(QTDIV*(1.D0+(KP1*XBMT)))
         BT1=QRDIV*(1.D0+(KP1*XBMW))/
     1       (QTDIV*(1.D0+(KP1*XBMT)))

C........MATRIX-VECTOR COEFFICIENTS
         IF (IC.GT.1) A(IC,IC-1) = -DNM1
         A(IC,IC) = D1N+D2N+((1.D0-AGA(IC))*D3N)+
     1             ((1.D0-AGA(IC))*D4N)+D5N 
         A(IC,NTKP1) = -DN*BT1
         B(IC)= DN*AT1

C.....END CSTR LOOP FOR AERATION BASIN
  110 CONTINUE

C.....SET UP COEFFICIENTS FOR SECONDARY CLARIFIER

C........FLOWRATES
         QSC=(QTDIV*DFLOAT(NAB))/DFLOAT(NSC)
         QINDIV=(QE1*DFLOAT(NPC))/DFLOAT(NSC)
         QTDSC=QINDIV-SQW

C........VOLATILIZATION FOR SEC. CLARIFIER 
         XKLAV2=KLAVL(CBSC,QVSC,SCBD,SCBV)

C........COMPONENTS OF COEFFICIENTS
         DN = QSC*(1.D0+(KP1*XBM(NTK)))
         DN1 = (QSC-QR-SQW)*(1+KP1*XSC)
         DN2 = (QR+SQW)*(1.D0+KP1*XBMW)
         DN3 = XKLAV2*SCBV

         IF (CBSC.EQ.-1) THEN
C...........INCLUDE GAS FROM SECONDARY EFFLUENT WEIR DROP
            IF (W3.EQ.-1) THEN
               RV3=WEIRV(WM3,Z3,QTDSC,WW3,QG3)
            ELSEIF (W3.EQ.0) THEN
               RV3=1.D0
            END IF
            FE3=1.D0-(1.D0/RV3)
            AG1=((DN3+(QSC*FE3))/
     1          ((QVSC*24.D0*60.D0)+(DN3/H)))/H
C...........CONVERT QVSC FROM L/MIN TO L/DAY
         ELSE
            AG1=0.D0
         END IF

C........MATRIX-VECTOR COEFFICIENTS
         A(NTKP1,NTK) = -DN
         A(NTKP1,NTKP1) = DN1+DN2+((1.D0-AG1)*DN3)
         B(NTKP1)=0.D0

C.....END COMPUTE MATRIX AND VECTOR COEFFICIENTS

C.....SOLVE MATRIX-VECTOR SYSTEM
C.....DECOMP AND SOLVE ARE TWO SUBROUTINES FOR SOLVING
C.....A SYSTEM OF N EQUATIONS, AX=B.
      CALL DECOMP(NTKP1,A,IPVT,DET)
      CALL SOLVE(NTKP1,A,IPVT,B)

C.....ASSIGN B TO CE2
      DO 200 IC=1,NTKP1
         CE2(IC)=B(IC)
  200 CONTINUE

C.....REMOVAL RATES FOR MULTIPLE CSTRS IN AERATION BASIN
      TRBIOS=0.D0
      TRSUFS=0.D0
      TRBULS=0.D0
      TRVLSA=0.D0
      DO 1100 IC=1,NTK
         RBIOSC(IC)=KB*XBM(IC)*CE2(IC)*ATV(IC)
     1              *DFLOAT(NAB)
         TRBIOS=TRBIOS+RBIOSC(IC)
         RSUFSC(IC)=XKLASU(IC)*CE2(IC)*(1.D0-AGA(IC))*ATV(IC)
     1              *DFLOAT(NAB)
         TRSUFS=TRSUFS+RSUFSC(IC)
         RBBLSC(IC)=QG(IC)*H*CE2(IC)*60.D0*24.D0
     1              *XBUB(IC)*DFLOAT(NAB)
         TRBULS=TRBULS+RBBLSC(IC)
         RVLSCA(IC)=XKLAV1*CE2(IC)*(1.D0-AGA(IC))*ATV(IC)
     1              *DFLOAT(NAB)
         TRVLSA=TRVLSA+RVLSCA(IC)
 1100 CONTINUE

C.....REMOVAL RATES FOR SECONDARY CLARIFIER
      RSOPSC=SQW*KP1*XBMW*CE2(NTKP1)*DFLOAT(NSC)
      RSOPSW=SQW*CE2(NTKP1)*DFLOAT(NSC)
      TRVLSB=XKLAV2*CE2(NTKP1)*(1.D0-AG1)*SCBV
     1        *DFLOAT(NSC)

C.....SECONDARY CLARIFIER WEIR
 
      IF (W3.EQ.-1) THEN
         RV3=WEIRV(WM3,Z3,QTDSC,WW3,QG3)
      ELSEIF (W3.EQ.0) THEN
         RV3=1.D0
      END IF
 
      FE3=1.D0-(1.D0/RV3)
      TCE2=CE2(NTKP1)*(1.D0-FE3)
      RVOLW3=QTDIV*CE2(NTKP1)*FE3*DFLOAT(NSC)
 
C
C.....STORE THIS VALUE FOR OUTPUT.
	XVALS_OUT(5) = XBM(NTK)
	XVALS_OUT(6) = XSC
	XVALS_OUT(7) = XSC
C.....END OF STORE THIS VALUE FOR OUTPUT.
C
      RETURN
      END


CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE REMOV 
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      INCLUDE 'blah.for'

      TIN=(Q*CO1*(1.D0+(KP1*X0)))
      PERVLP=100.D0*RVOLPRI/TIN
      PERSPA=100.D0*RSORPI/TIN
      PERSPC=100.D0*RSORPW/TIN
      PERSPB=100.D0*RSOPSC/TIN
      PERSPD=100.D0*RSOPSW/TIN
      PERBIO=100.D0*TRBIOS/TIN
      PERBUB=100.D0*TRBULS/TIN
      PERSUF=100.D0*TRSUFS/TIN
      PERVLA=100.D0*TRVLSA/TIN
      PERVLB=100.D0*TRVLSB/TIN
      PRVLGC=100.D0*RVOLAG/TIN
      PERW1=100.D0*RVOLW1/TIN
      PERAGC=100.D0*RSTRPG/TIN
      PERW2=100.D0*RVOLW2/TIN
      PERW3=100.D0*RVOLW3/TIN

      PESTAB=PERSUF+PERBUB
      PEPSTP=PERBUB+PERSUF+PERW1+PERW2+PERW3+PERAGC
      PEPVOL=PERVLA+PERVLB+PERVLP+PRVLGC          
      PERSP=PERSPA+PERSPB
      PERSPW=PERSPC+PERSPD

      TRSTRP=TRBULS+TRSUFS+RVOLW1+RVOLW2+RVOLW3+RSTRPG
      TRSVOL=TRVLSA+TRVLSB+RVOLPRI+RVOLAG
      TRORPS=RSOPSC+RSORPI
      TRORPW=RSORPW+RSOPSW

      TPERX=PERSP+PEPSTP+PERBIO+PEPVOL+PERSPW


      TEFF=TCE2*(Q-(QW1*DFLOAT(NPC))-(SQW*DFLOAT(NSC)))*
	+     (1+KP1*XSC)
      TPER = (1.D0-(TEFF/TIN))*100.D0	

C.....CONVERT REMOVALS TO KG/DAY
      RVOLW1=RVOLW1/1.D9
      RVOLAG=RVOLAG/1.D9
      RSTRPG=RSTRPG/1.D9
      RVOLPRI=RVOLPRI/1.D9
      RSORPI=RSORPI/1.D9
      RSORPW=RSORPW/1.D9
      RSOPSC=RSOPSC/1.D9
      RSOPSW=RSOPSW/1.D9
      RVOLW2=RVOLW2/1.D9
      TRVLSB=TRVLSB/1.D9
      RVOLW3=RVOLW3/1.D9
      TRORPS=TRORPS/1.D9
      TRORPW=TRORPW/1.D9
      TRBIOS=TRBIOS/1.D9
      TRSTAB=(TRSUFS+TRBULS)/1.D9
      TRVLSA=TRVLSA/1.D9
      TRSTRP=TRSTRP/1.D9
      TRSVOL=TRSVOL/1.D9
      TIN=TIN/1.D9
      TEFF=TEFF/1.D9


      RETURN
      END

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE OUTDAT
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      INCLUDE 'blah.for'

      OPEN (UNIT=20, FILE='output.txt')

      WRITE (20, *) 
      WRITE (20,2007) ' CONTAMINANT NAME ',NAME 
      WRITE (20, *) 
      WRITE (20,2001) 'BIODEGRADATION CONSTANT ',KB
      WRITE (20,2001) 'HENRYS CONSTANT DIMENSIONLESS ',H
      WRITE (20,2001) 'LOG KOW ',LOGKOW
      WRITE (20, *) 
      WRITE (20,2005) 'PERCENT REMOVAL BY BIODEGRADATION ',PERBIO           
      WRITE (20,2005) 'PERCENT REMOVAL BY SOLID WASTAGE ',PERSP            
      WRITE (20,2005) 'PERCENT REMOVAL BY LIQUID WASTAGE ',PERSPW            
      WRITE (20,2005) 'PERCENT REMOVAL BY STRIPPING ',PEPSTP                   
      WRITE (20,2005) 'PERCENT REMOVAL BY VOLATILIZATION ',PEPVOL
      WRITE (20,2005) 'TOTAL PERCENT REMOVED (by concns) ',TPER
      WRITE (20,2005) 'TOTAL PERCENT REMOVED (by removals) ',TPERX
      WRITE (20, *) 
      WRITE (20,2002) 'TOTAL AMOUNT IN INFLUENT ',TIN,' kg/day'
      WRITE (20,2002) 'TOTAL AMOUNT BIODEGRADED ',TRBIOS,' kg/day'
      WRITE (20,2002) 'TOTAL AMOUNT WASTED AS SOLID ',TRORPS,' kg/day'
      WRITE (20,2002) 'TOTAL AMOUNT WASTED AS LIQ. ',TRORPW,' kg/day'
      WRITE (20,2002) 'TOTAL AMOUNT STRIPPED ',TRSTRP,' kg/day'
      WRITE (20,2002) 'TOTAL AMOUNT VOLATILIZED ',TRSVOL,' kg/day'
      WRITE (20,2002) 'TOTAL AMOUNT IN EFFLUENT ',TEFF,' kg/day'
      WRITE (20, *) 
      WRITE (20,*) ' RESULTS BY MECHANISM AND UNIT OPERATION'
      WRITE (20, *) 
      WRITE (20,2006) ' Influent Liquid Concn',CO1,' ug/L'
      IF (W1.EQ.-1) THEN
         WRITE (20, *) 
         WRITE (20, *) 'INFLUENT WEIR'
         WRITE (20, *) 
         WRITE (20,2004) RVOLW1,' kg/day REMOVED BY STRIP.',  
     +'  ( ',PERW1,'% OF TOTAL)'
         WRITE (20,2006) ' Effluent Liquid Concn',CA,' ug/L'
      END IF
      IF (GC.EQ.-1) THEN
         WRITE (20, *) 
         WRITE (20, *) 'AERATED GRIT CHAMBER'
         WRITE (20, *) 
         WRITE (20,2004) RVOLAG,' kg/day REMOVED BY VOL.',  
     +'  (',PRVLGC,'% OF TOTAL) '
         WRITE (20,2004) RSTRPG,' kg/day REMOVED BY STRIP.',
     +'  (',PERAGC,'% OF TOTAL) '
         WRITE (20,2006) ' Effluent Liquid Concn',CB,' ug/L'
      END IF
         WRITE (20, *) 
      WRITE (20, *) 'PRIMARY CLARIFIER'
         WRITE (20, *) 
      WRITE (20,2004) RVOLPRI,' kg/day REMOVED BY VOL.',
     +'  (',PERVLP,'% OF TOTAL) '
      WRITE (20,2004) RSORPI,' kg/day REMOVED BY SOLID WAST.',
     +'  ( ',PERSPA,'% OF TOTAL) '
      WRITE (20,2004) RSORPW,' kg/day REMOVED BY LIQ. WAST.',
     +'  ( ',PERSPC,'% OF TOTAL) '
      WRITE (20,2006) ' Effluent Liquid Concn',CC,' ug/L'
      IF (W2.EQ.-1) THEN
         WRITE (20, *) 
         WRITE (20, *) 'PRIMARY CLARIFIER WEIR'
         WRITE (20, *) 
         WRITE (20,2004) RVOLW2,' kg/day REMOVED BY STRIP.',
     +'  (',PERW2,'% OF TOTAL) '
         WRITE (20,2006) ' Effluent Liquid Concn',CD,' ug/L'
      END IF
         WRITE (20, *) 
      WRITE (20,2003) 'AERATION BASIN MODELED AS',NTK,' CSTRS'
         WRITE (20, *) 
      WRITE (20,2004) TRVLSA,' kg/day REMOVED BY VOL.',
     +'  (',PERVLA,'% OF TOTAL) '
      WRITE (20,2004) TRSTAB,' kg/day REMOVED BY STRIP.',
     +'  (',PESTAB,'% OF TOTAL) '
      WRITE (20,2004) TRBIOS,' kg/day REMOVED BY BIODEG.',
     +'  (',PERBIO,'% OF TOTAL) '
      WRITE (20,2006) ' Effluent Liquid Concn',CE2(NTK),' ug/L'
      WRITE (20, *) 
      WRITE (20, *) 'SECONDARY CLARIFIER'
         WRITE (20, *) 
      WRITE (20,2004) TRVLSB,' kg/day REMOVED BY VOL.',
     +'  (',PERVLB,'% OF TOTAL) '
      WRITE (20,2004) RSOPSC,' kg/day REMOVED BY SOLID WAST.',
     +'  (',PERSPB,'% OF TOTAL) '
      WRITE (20,2004) RSOPSW,' kg/day REMOVED BY LIQ. WAST.',
     +'  (',PERSPD,'% OF TOTAL) '
      WRITE (20,2006) ' Effluent Liquid Concn',CE2(NTKP1),' ug/L'
      WRITE (20, *) 
      IF (W3.EQ.-1) THEN
         WRITE (20, *) 'SECONDARY CLARIFIER WEIR'
         WRITE (20, *) 
         WRITE (20,2004) RVOLW3,' kg/day REMOVED BY STRIP.',
     +'  (',PERW3,'% OF TOTAL) '
         WRITE (20,2006) ' Effluent Liquid Concn',TCE2,' ug/L'
      END IF
	CLOSE(UNIT=20)
       
 2001 FORMAT(' ',A40,1P,E12.5)
 2002 FORMAT(' ',A40,1P,E12.5,' ',A10)
 2003 FORMAT(' ',A25,I3,' ',A6)
 2004 FORMAT(' ',1P,E12.5,A30,A3,0P,F8.4,' ',A12)
 2005 FORMAT(' ',A40,F8.4)
 2006 FORMAT(' ',A22,1P,E9.2,A5)
 2007 FORMAT(' ',A20,A20)
        RETURN
        END

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE VBOUTPUT
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      INCLUDE 'blah.for'

      OPEN (UNIT=20, FILE='voc.out')

C PERCENT REMOVAL BY STRIPPING
      WRITE (20,3001) PEPSTP                   
C PERCENT REMOVAL BY VOLATILIZATION
      WRITE (20,3001) PEPVOL
C PERCENT REMOVAL BY SOLID WASTAGE
      WRITE (20,3001) PERSP            
C PERCENT REMOVAL BY LIQUID WASTAGE
      WRITE (20,3001) PERSPW            
C PERCENT REMOVAL BY BIODEGRADATION
      WRITE (20,3001) PERBIO           
C TOTAL PERCENT REMOVED (by removals)
      WRITE (20,3001) TPERX

C TOTAL AMOUNT STRIPPED (kg/day)
      WRITE (20,3001) TRSTRP
C TOTAL AMOUNT VOLATILIZED (kg/day)
      WRITE (20,3001) TRSVOL
C TOTAL AMOUNT WASTED AS SOLID (kg/day)
      WRITE (20,3001) TRORPS
C TOTAL AMOUNT WASTED AS LIQUID (kg/day)
      WRITE (20,3001) TRORPW
C TOTAL AMOUNT BIODEGRADED (kg/day)
      WRITE (20,3001) TRBIOS
C TOTAL AMOUNT IN INFLUENT (kg/day)
      WRITE (20,3001) TIN
C TOTAL AMOUNT IN EFFLUENT (kg/day)
      WRITE (20,3001) TEFF


C ****************** INFLUENT WEIR

      IF (W1.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CA
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RVOLW1
          WRITE (20,3001) PERW1
      END IF

C ****************** AERATED GRIT CHAMBER

      IF (GC.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CB
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RSTRPG
          WRITE (20,3001) PERAGC
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) RVOLAG
          WRITE (20,3001) PRVLGC
      END IF

C ****************** PRIMARY CLARIFIER

C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CC
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) RVOLPRI
          WRITE (20,3001) PERVLP
C         kg/day REMOVED BY SOLID WASTE | % OF TOTAL
          WRITE (20,3001) RSORPI
          WRITE (20,3001) PERSPA
C         kg/day REMOVED BY LIQ. WASTE | % OF TOTAL
          WRITE (20,3001) RSORPW
          WRITE (20,3001) PERSPC

C ****************** PRIMARY CLARIFIER WEIR

      IF (W2.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CD
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RVOLW2
          WRITE (20,3001) PERW2
      END IF

C ****************** AERATION BASIN

C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CE2(NTK)
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) TRSTAB 
          WRITE (20,3001) PESTAB
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) TRVLSA
          WRITE (20,3001) PERVLA
C         kg/day REMOVED BY BIODEG | % OF TOTAL
          WRITE (20,3001) TRBIOS
          WRITE (20,3001) PERBIO

C ****************** SECONDARY CLARIFIER

C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) CE2(NTKP1)
C         kg/day REMOVED BY VOL | % OF TOTAL
          WRITE (20,3001) TRVLSB
          WRITE (20,3001) PERVLB
C         kg/day REMOVED BY SOLID WASTE | % OF TOTAL
          WRITE (20,3001) RSOPSC
          WRITE (20,3001) PERSPB
C         kg/day REMOVED BY LIQ. WASTE | % OF TOTAL
          WRITE (20,3001) RSOPSW
          WRITE (20,3001) PERSPD

C ****************** SECONDARY CLARIFIER WEIR

      IF (W3.EQ.-1) THEN
C         Effluent Liquid Concn (ug/L)
          WRITE (20,3001) TCE2
C         kg/day REMOVED BY STRIP | % OF TOTAL
          WRITE (20,3001) RVOLW3
          WRITE (20,3001) PERW3
      END IF

C
C.....STORE NEW VALUE FOR OUTPUT (1999-SEP-22).
      WRITE (20,3001) KP1
	DO I=1,7
	  WRITE (20,3001) XVALS_OUT(I) 
      ENDDO
C.....END STORE NEW VALUE FOR OUTPUT (1999-SEP-22).
C


 	CLOSE(UNIT=20)
      
 3001 FORMAT(' ',1P, E20.12)

        RETURN
        END




CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                          C
C    FUNCTION KLAVL                                        C
C                                                          C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC

      DOUBLE PRECISION FUNCTION KLAVL(NCV,QVNT,DEPTH,VOLUME)
      INCLUDE 'blah.for'
      INTEGER NCV

C.....FRICTION VELOCITY (IN M/S)
      IF (NCV.EQ.0) THEN
         FV=((6.1D0+0.63D0*WNDVRI)**0.5D0)*WNDVRI
      ELSEIF (NCV.EQ.-1) THEN
         FV=QVNT/(VOLUME/DEPTH)
      ENDIF

C.....VOLATILIZATION MT COEFF.
C.....MACKAY AND YEUN CORRELATION
      IF (PVM.EQ.1) THEN
         IF (FV.GT.0.3) THEN
C...........FOR AIR SIDE FRICTION VELOCITIES > 0.3 M/S
            KL=(1.0D-6)+(34.1D-4)*FV*((SCLPRI)**(-0.5))
         ELSE
C...........FOR AIR SIDE FRICTION VELOCITIES < 0.3 M/S
            KL=(1.0D-6)+(144.0D-4)*(FV**2.2D0)
     1         *(SCLPRI**(-0.5D0))
         END IF
         KG=(1.D-3)+(46.2D-3*FV*(SCGPRI**(-0.67D0)))
         KKL=3600.D0*24.D0*(1.D0/((1.D0/KL)+(1.D0/(H*KG))))
C........CONVERT KKL FROM M/SEC TO M/DAY
         KLAVL=KKL/DEPTH
 
C.....DOBBS/COHEN CORRELATION
      ELSEIF (PVM.EQ.2) THEN
         Z0=(FV*100.D0)/500.D0
         AIRKV=AIRVIS*1.D+04/AIRDEN
C........CONVERT AIRKV FROM M2/S TO CM2/S
         RESTAR=Z0*(FV*100.D0)/AIRKV
         VOCTOL=6.88D-6
         DIFRAT=VOCDIF/VOCTOL
         KLA=((11.4*(RESTAR**0.195D0))-5.D0)*
     1       DIFRAT*(1.D0/DEPTH)*(24.D0/100.D0)
         KGA=700.D0*((18.D0/MW)**0.25D0)*FV*
     1       (24.D0/(100.D0*DEPTH))
         KKLA=1.D0/((1.D0/KLA)+(1.D0/(H*KGA)))
         KLAVL=KKLA
 
      END IF

      RETURN
      END
 

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                          C
C    FUNCTION BUBPAR                                       C
C                                                          C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC

      DOUBLE PRECISION FUNCTION BUBPAR(BSOTR,QGAS,VOLUME,BDEP)
      INCLUDE 'blah.for'

C.....CALCULATE EFFECTIVE DEPTH AND CSTAR INFINITY
C.....PS=STANDARD PRESSURE IN KPA
      PS=101.33D0
      DEEQ=0.35D0*BDEP
      CSI=CSAT*((H2OSW*DEEQ)+PB-PV)/(PS-PV)

      PHIC=(AIRDEN*QGAS*HOXY*60.D0*8.314D0*(T+273.D0))
     1    /(AIRMW*VOLUME*(PB+(H2OSW*DEEQ)))
C.....CONVERT PHIC FROM 1/MIN TO 1/HR
C.....CONVERT T FROM CELCIUS TO KELVIN 

      BLKOXY=ALPHA*BSOTR*1.D6*(1.024D0**(T-20.D0))
     1      /(VOLUME*CSI)
C.....CONVERT CSI FROM MG/L TO KG/L TO MATCH BSOTR UNITS

      KLATRU=BLKOXY/(1.D0+(BLKOXY/(2.D0*PHIC)))
      BUKVOC=KLATRU*((VOCDIF/OXYDIF)**0.5D0)
      PHI=(BUKVOC*VOLUME)/(H*QGAS*60.D0)
C.....CONVERT BUKVOC FROM 1/HR TO 1/MIN TO MATCH QGAS UNITS


      IF (PHI.GT.100.D0) PHI=100.D0
      BUBPAR=1.D0-DEXP(-PHI)

      RETURN
      END
 
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                          C
C    FUNCTION KLASUF                                       C
C                                                          C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC

      DOUBLE PRECISION FUNCTION KLASUF(VOLUME,BDEP)
      INCLUDE 'blah.for'

C.....CALCULATE EFFECTIVE DEPTH AND CSTAR INFINITY
C.....PS=STANDARD PRESSURE IN KPA
      PS=101.33D0
      DEEQ=0.07D0*BDEP
      CSI=CSAT*((H2OSW*DEEQ)+PB-PV)/(PS-PV)

      SULOXY=ALPHA*SUFSOR*1.D6*(1.024D0**(T-20.D0))
     1            /(VOLUME*CSI)
C.....CONVERT CSI FROM MG/L TO KG/L TO MATCH SUFSOR UNITS
      SULVOC=SULOXY*(VOCDIF/OXYDIF)**0.5D0
      KLASUF=3600.D0*24.D0*(1.D0/((1.D0/SULVOC)+
     1       (1.D0/(40.D0*SULVOC*H))))
C.....CONVERT KLASUF FROM 1/SEC TO 1/DAY

      RETURN
      END
 

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                          C
C    FUNCTION WEIRV                                        C
C                                                          C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
      DOUBLE PRECISION FUNCTION WEIRV(MODW,Z,QIN,WIDTH,QGW)
      INCLUDE 'blah.for'
      INTEGER MODW

      QL=QIN/(WIDTH*24.D0*1000.D0)
C.....QIN is (L/day), WIDTH is (m), QGW is m3/(m h)
      QGQL=QGW/QL

      IF (MODW.EQ.1) THEN
C.....NAPPE MODEL
         QL=QL/3600.D0
C........CONVERT QL FROM M3/(M H) TO M3/(M S)
         LROXY=0.042D0*(Z**0.872D0)*(QL**0.509D0)
         LROXT=(1.D0+0.0168D0*(T-20.D0))*LROXY
         LRVOC=((VOCDIF/OXYDIF)**0.5D0)*LROXT
         WEIRV=DEXP(LRVOC)
      ELSEIF (MODW.EQ.2) THEN
C.....POOL MODEL
         LROXY=0.042D0*(Z**0.872D0)*(QL**0.509D0)
         LROXT=(1.D0+0.0168D0*(T-20.D0))*LROXY
         RO=DEXP(LROXT)
         WEIRV=1.D0+(QGQL)*H*
     1      (1.D0-DEXP(((VOCDIF/OXYDIF)**0.5D0)*
     2      (HOXY/H)*DLOG(1.D0-(RO-1.D0)/(HOXY*QGQL))))
       END IF
       
       RETURN
       END
 

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
C          SUBROUTINE DECOMP                                         
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC


C PROGRAM FOR SOLVING A LINEAR SYSTEM OF N EQUATIONS
C AX=B

C note that the solution, X,
C is written over the b vector
    
      SUBROUTINE DECOMP (N,A,IPVT,DET)
     
      IMPLICIT REAL*8 (A-M,O-Z) 
      INTEGER N,IPVT(10)
      REAL *8 A(10,10),DET
      REAL *8 P,T
      INTEGER NM1,I,J,K,KPL1,M


      DET = 1.D0
      IPVT(N) = 1
      NM1 = N-1

      DO 60 K = 1,NM1
      KPL1 = K+1

C     FIND PIVOT P

      M = K
      DO 10 I = KPL1,N
   10 IF (DABS(A(I,K)).GT.DABS(A(M,K))) M = I
      IPVT(K) = M
      IF (M.NE.K) IPVT(N) = -IPVT(N)
      P = A(M,K)
      A(M,K) = A(K,K)
      A(K,K) = P
      DET = DET*P
      IF (P.EQ.0.D0) GO TO 60

C     COMPUTE MULTIPLIERS

   20 DO 30 I = KPL1,N
   30 A(I,K) = -A(I,K)/P

C     INTERCHANGE AND ELIMINATE BY COLUMNS

      DO 50 J = KPL1,N
      T = A(M,J)
      A(M,J) = A(K,J)
      A(K,J) = T
      IF (T.EQ.0.D0) GO TO 50
      DO 40 I = KPL1,N
      A(I,J) = A(I,J) + A(I,K)*T
   40 CONTINUE
   50 CONTINUE
   60 CONTINUE

   70 DET = DET*A(N,N)*DFLOAT(IPVT(N))

      RETURN
      END

CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
C         SUBROUTINE SOLVE                                         
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC

      SUBROUTINE SOLVE (N,A,IPVT,B)

      IMPLICIT REAL*8 (A-M,O-Z) 
      INTEGER IPVT(10)
      REAL *8 A(10,10),B(10)
      INTEGER N,NM1,K,KB,KPL1,KM1,M,I
      REAL *8 S

C     FORWARD ELIMINATION

      NM1 = N-1
      DO 10 K = 1,NM1
      KPL1 = K+1
      M = IPVT(K)
      S = B(M)
      B(M) = B(K)
      B(K) = S
      DO 10 I = KPL1,N
   10 B(I) = B(I)+A(I,K)*S

C     BACK SUBSTITUTION

      DO 20 KB = 1,NM1
      KM1 = N-KB
      K = KM1+1
      B(K) = B(K)/A(K,K)
      S = -B(K)
      DO 20 I = 1,KM1
   20 B(I) = B(I)+A(I,K)*S
   30 B(1) = B(1)/A(1,1)

      RETURN
      END
