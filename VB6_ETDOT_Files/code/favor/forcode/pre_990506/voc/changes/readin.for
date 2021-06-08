
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                   C
C                                                                   C
        SUBROUTINE READIN
C                                                                   C
C                                                                   C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
        INCLUDE 'blah.for'
        INTEGER IC

        OPEN (UNIT=10, FILE='input1.dat', STATUS='OLD')

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
        READ (10,27) NAME
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
C SOTR For Bubble Aeration in Aeration Basin (kg/hr),
        READ (10,20) ABBSOT
C Secondary Aeration Mechanism (1=SURFACE 3=DIFFUSED BUBBLE)
	  READ (10,27) SAM 


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



C *************** UNKNOWN WHERE THIS COMES FROM
        READ (10,*) 
        READ (10,*) 
        READ (10,*) 

C SOTR For Surface Aeration (kg/hr)
        READ (10,20) SUFSOR

        CLOSE(UNIT=10)
       
C *************** AERATION BASINS CSTRs


        OPEN (UNIT=14, FILE='input2.dat', STATUS='OLD')
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
27  	FORMAT (/ I2)       

        RETURN
        END


