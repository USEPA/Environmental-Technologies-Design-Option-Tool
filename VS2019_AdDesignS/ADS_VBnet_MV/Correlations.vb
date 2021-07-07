Option Strict Off
Option Explicit On
Module Correlations
	
	
	
	'C------ CALCULATE WATER DENSITY (kg/m^3) VIA MTU CORRELATION.
	'INPUTS:
	'    - TEMPERATURE as Input_Temp, degK.
	'RETURNS:
	'    - WATER DENSITY, g/cm^3.
	Function KineticCorr_WaterDensity(ByRef Input_Temp As Object) As Double
		Dim TA As Double
		Dim RetVal As Double
		Dim A1 As Double
		Dim A2 As Double
		Dim A3 As Double
		Dim A4 As Double
		Dim A5 As Double
		Dim XAVG As Double
		Dim FAVG As Double
		A1 = -1.4176800403
		A2 = 8.976651524
		A3 = -12.275501969
		A4 = 7.4584410413
		A5 = -1.738491605
		XAVG = 324.65
		FAVG = 0.98396
		'NOTE: Input_Temp IS IN UNITS OF DEGK.
		'UPGRADE_WARNING: Couldn't resolve default property of object Input_Temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TA = Input_Temp / XAVG
		RetVal = (A1 + A2 * TA + A3 * TA ^ 2# + A4 * TA ^ 3# + A5 * TA ^ 4#) * FAVG * 1000#
		'(NOTE 1000# FACTOR IS TO CONVERT g/cm^3 TO kg/m^3.)
		KineticCorr_WaterDensity = RetVal
	End Function
	'C------ CALCULATE WATER VISCOSITY (g/cm-s) VIA YAWS CORRELATION.
	'INPUTS:
	'    - TEMPERATURE as Input_Temp, degK.
	'RETURNS:
	'    - WATER VISCOSITY, kg/m-s
	'      (THIS WAS INCORRECTLY LABELLED AS g/cm-s BEFORE 1/6/99).
	Function KineticCorr_WaterViscosity(ByRef Input_Temp As Double) As Double
		Dim TB As Double
		Dim RetVal As Double
		'NOTE: Input_Temp IS IN UNITS OF DEGK.
		TB = Input_Temp
		RetVal = System.Math.Exp((-24.71) + (4209#) / TB + (0.04527) * TB - (0.00003376) * TB ^ 2#)
		RetVal = RetVal / 1000#
		KineticCorr_WaterViscosity = RetVal
	End Function
	'C------ CALCULATE AIR VISCOSITY (g/cm-s) VIA HOKANSON CODE.
	'INPUTS:
	'    - TEMPERATURE as Input_Temp, degK.
	'RETURNS:
	'    - AIR VISCOSITY, kg/m-s
	'      (THIS WAS INCORRECTLY LABELLED AS g/cm-s BEFORE 1/6/99).
	Function KineticCorr_AirViscosity(ByRef Input_Temp As Double) As Double
		'NOTE: Input_Temp IS IN UNITS OF DEGK.
		KineticCorr_AirViscosity = 0.00000017 * (Input_Temp ^ 0.818)
	End Function
	'C------ CALCULATE AIR DENSITY (kg/m^3) VIA HOKANSON CODE.
	'INPUTS:
	'    - TEMPERATURE as Input_Temp, degK.
	'    - PRESSURE as Input_Pres, atm.
	'RETURNS:
	'    - AIR DENSITY, kg/m^3.
	Function KineticCorr_AirDensity(ByRef Input_Temp As Double, ByRef Input_Pres As Double) As Double
		Dim MWAVG As Double
		Dim r As Double
		Dim ThisVal As Double
		'NOTE: Input_Temp IS IN UNITS OF DEGK.
		'NOTE: Input_Pres IS IN UNITS OF ATM.
		MWAVG = 28.95 'UNITS OF G/GMOL.
		r = 0.08205 'UNITS OF (L-ATM)/(GMOL-K).
		'
		' ON THE NEXT LINE, THE UNITS ARE gmol/L =
		' (atm/((L-atm)/(gmol-K))/(K)) =
		' (atm*gmol-K)/(L-atm-K) = gmol/L (CHECKS).
		ThisVal = (Input_Pres / r / Input_Temp)
		'
		' ON THE NEXT LINE, THE UNITS ARE g/L = (gmol/L)*(g/gmol) (CHECKS).
		ThisVal = ThisVal * MWAVG
		'
		' ON THE NEXT LINE, THE UNITS ARE kg/m^3 = g/L (CHECKS).
		ThisVal = ThisVal * 1#
		'
		' RETURN THE VALUE.
		KineticCorr_AirDensity = ThisVal
		'KineticCorr_AirDensity = _
		'(1# / 1000#) * ((MWAVG) * (Input_Pres)) / ((r) * (Input_Temp))
	End Function
	
	
	Sub Update_FluidDensity(ByRef Temperature As Double, ByRef Pressure As Double, ByRef Density As Double)
		Dim Dummy As Double
		If (State_Check_Water(1)) Then
			If (Bed.Phase = 0) Then
				'Call H2ODens(dummy, Temperature + 273.15)
				Dummy = KineticCorr_WaterDensity(Temperature + 273.15)
			Else
				'Call AIRDens(dummy, Temperature + 273.15, Pressure)
				Dummy = KineticCorr_AirDensity(Temperature + 273.15, Pressure)
			End If
			Density = Dummy / 1000#
			'(NOTE 1/1000 FACTOR IS TO CONVERT kg/m^3 TO g/cm^3.)
		End If
		'Note: If density correlation not in use, Density
		'is returned unchanged.
	End Sub
	Sub Update_FluidViscosity(ByRef Temperature As Double, ByRef Viscosity As Double)
		Dim Dummy As Double
		If (State_Check_Water(2)) Then
			If (Bed.Phase = 0) Then
				'Call H2OVisc(dummy, Temperature + 273.15)
				Dummy = KineticCorr_WaterViscosity(Temperature + 273.15)
			Else
				'Call AirVisc(dummy, Temperature + 273.15)
				Dummy = KineticCorr_AirViscosity(Temperature + 273.15)
			End If
			Viscosity = Dummy * 10#
			'(NOTE 10# FACTOR IS TO CONVERT kg/m-s TO g/cm-s.)
		End If
		'Note: If viscosity correlation not in use, Viscosity
		'is returned unchanged.
	End Sub
	
	
	Function Diffg(ByRef i As Short) As Double
		'Wilke-Lee Correlation
		Dim WTAIR, Press As Double
		Dim VOLA, RCOM, RAIR, RAB As Double
		Dim FE, EAB, E, DiffgT As Double
		WTAIR = 28.964
		RAIR = 0.3711
		'Press = Bed.Pressure * 101.3
		'Converting P from atm to Pa
		Press = Bed.Pressure * 101325
		'Converting Molar Volume @ NBP from cm3/mol to m3/kmol
		VOLA = Component(i).MolarVolume / 1000#
		RCOM = 1.18 * VOLA ^ 0.333333333333
		RAB = (RAIR + RCOM) / 2#
		'Energy of molecular attraction
		'78.6 is Eair and Ecompound is calculated using : e/k = 1.21 * Tb (Tb = NBP)
		' Actually, EAB is = to epsilonAB over k
		EAB = System.Math.Sqrt(78.6 * 1.21 * (Component(i).BP + 273.15))
		'Collision function
		E = System.Math.Log((Bed.Temperature + 273.15) / EAB) / System.Math.Log(10#)
		FE = 10 ^ (-0.14329 - 0.48343 * E + 0.1939 * E ^ 2 + 0.13612 * E ^ 3 - 0.20578 * E ^ 4 + 0.083899 * E ^ 5 - 0.011491 * E ^ 6)
		DiffgT = (0.0001 * (1.084 - 0.249 * System.Math.Sqrt(1# / WTAIR + 1# / Component(i).MW)) * (Bed.Temperature + 273.15) ^ 1.5 * System.Math.Sqrt(1# / WTAIR + 1# / Component(i).MW)) / (Press * RAB ^ 2 * FE)
		
		
		
		'C
		'C CALCULATION OF GAS DIFFUSIVITY USING WILKE-LEE MODIFICATION OF
		'C THE HIRSCHFELDER-BIRD-SPOTZ CORRELATION (m^2/sec)
		'C
		'       MSAC(I) = (1.18D0*(VB(I)/1000.0d0)**(1.0D0/3.0D0)+
		'     1  707        0.3711D0)/2.0D0
		'     ROOT(i) = SQRT(1# / XWT(i) + 1# / 28.964)
		'        print *, 'xaxa - ROOT(i) = ', ROOT(I)
		'
		'       GDIF(I) = (1.084D-4 - 2.49D-5 * ROOT(I)) * TEMP**1.5D0 *
		'     &               ROOT(I) / (PAS * MSAC(I)**2.0D0 * CF(I))
		
		
		'Converting from m2/s to cm2/s
		Diffg = DiffgT * 100# * 100#
	End Function
	Function Diffl(ByRef i As Short) As Double
		'Diffl is in cm2/s
		Diffl = 0.0001326 / (100# * Bed.WaterViscosity) ^ 1.14 / (Component(i).MolarVolume) ^ 0.589
	End Function
	
	
	Function Dp(ByRef i As Short) As Double
		'Dp is in cm2/s
		If (Bed.Phase = 0) Then
			Dp = Diffl(i) / Component(i).Tortuosity
		Else
			Dp = Diffg(i) / Component(i).Tortuosity
		End If
	End Function
	Function Ds(ByRef i As Short) As Double
		'Ds is in cm2/s
		'Dl and Dg are in cm2/s
		'Carbon.Density is converted from g/cm3 -> g/l  by multipling by 1000
		'C0 is in mg/l
		'q0 is in mg/g
		If ((Component(i).Use_Tortuosity_Correlation) And (Component(i).Constant_Tortuosity)) Then
			'---- Force Ds to 1.0e-30
			Ds = 1E-30
		Else
			If (Bed.Phase = 0) Then
				Ds = Carbon.Porosity * Diffl(i) * Component(i).InitialConcentration * Component(i).SPDFR / (1000# * Carbon.Density * Component(i).Tortuosity * (Component(i).InitialConcentration) ^ Component(i).Use_OneOverN * Component(i).Use_K)
			Else
				Ds = Carbon.Porosity * Diffg(i) * Component(i).InitialConcentration * Component(i).SPDFR / (1000# * Carbon.Density * Component(i).Tortuosity * (Component(i).InitialConcentration) ^ Component(i).Use_OneOverN * Component(i).Use_K)
			End If
		End If
	End Function
	
	
	Function kf(ByRef i As Short) As Double
		Dim Re As Double 'Reynolds number
		Dim Vs As Double 'Superficial velocity
		Dim Vi As Double 'Interstitial velocity
		Dim EBED As Double 'Bed Porosity
		Dim ss As Double 'Schmidt number
		Dim D As Double
		If Error_In_Kinetic_Calculation Then
			kf = 0.00125
			Exit Function
		End If
		On Error GoTo Error_KF
		If (Bed.Phase = 0) Then
			'>>>>>> >>>   LIQUID PHASE   <<< <<<<<<
			'EBED IS DIMENSIONLESS.
			EBED = 1 - Bed.Weight / Bed.Diameter ^ 2 / PI * 4 / Bed.length / Carbon.Density / 1000#
			'Vs is in m/s
			Vs = Bed.Flowrate * 4# / Bed.Diameter ^ 2 / PI
			Vi = Vs / EBED
			'The Reynolds number is the INTERSTITIAL Reynolds Number
			'The viscosity is converted from g/cm.s to kg/m.s by dividing its value by 10.0
			Re = 2# * Carbon.ParticleRadius * Vi * Bed.WaterDensity * 1000# / (Bed.WaterViscosity / 10#)
			ss = SC(i)
			'KF is in cm/s
			'Diffg is in cm2/s, Carbon.ParticleRadius is in m -> Kf is in cm/s
			'The Gnielinski correlation is used instead of the following one.
			'kf = 2.4 * Vs * 100 / Re ^ .66 / SS ^ .58
			kf = (1# + 1.5 * (1 - EBED)) * Carbon.ShapeFactor * Diffl(i) / (2 * Carbon.ParticleRadius * 100#) * (2# + 0.644 * Re ^ 0.5 * ss ^ (1 / 3))
		Else
			'>>>>>> >>>   GAS PHASE   <<< <<<<<<
			'EBED IS DIMENSIONLESS.
			EBED = 1 - Bed.Weight / Bed.Diameter ^ 2 / PI * 4 / Bed.length / Carbon.Density / 1000#
			'Vs is in m/s
			Vs = Bed.Flowrate * 4# / Bed.Diameter ^ 2 / PI
			Vi = Vs / EBED
			If (USE_GASPHASE_WAKAO_AND_FUNAZUKRI) Then
				'USE THE WAKAO-FUNAZUKRI CORRELATION.
				'---- Old code: Wakao-Funzaki
				'The Reynolds number is the SUPERFICIAL Reynolds Number
				'The viscosity is converted from g/cm.s to kg/m.s by dividing its value by 10.0
				Re = 2# * Carbon.ParticleRadius * Vs * Bed.WaterDensity * 1000# / (Bed.WaterViscosity / 10#)
				ss = SC(i)
				'Diffg is in cm2/s, Radius is in m -> Kf is in cm/s
				kf = Diffg(i) / (2# * Carbon.ParticleRadius * 100#) * Carbon.ShapeFactor * (2# + 1.1 * Re ^ 0.6 * ss ^ (1# / 3#))
				'---- End of Old code
			Else
				'USE THE GNIELINSKI CORRELATION.
				'---- New kf for gas phase (Gnielinski Correlation) added by EJO on 11/1/96
				'The Reynolds number is the INTERSTITIAL Reynolds Number
				'The viscosity is converted from g/cm.s to kg/m.s by dividing its value by 10.0
				'Note: for gas phase, variables Bed.WaterDensity and Bed.WaterViscosity store
				'         Air Density and Air Viscosity, respectively
				Re = 2# * Carbon.ParticleRadius * Vi * Bed.WaterDensity * 1000# / (Bed.WaterViscosity / 10#)
				ss = SC(i)
				'Diffg is in cm2/s, Radius is in m -> Kf is in cm/s
				kf = (1# + 1.5 * (1# - EBED)) * Carbon.ShapeFactor * Diffg(i) / (2# * Carbon.ParticleRadius * 100#) * (2# + 0.644 * Re ^ 0.5 * ss ^ (1# / 3#))
			End If
		End If
		Exit Function
Error_KF: 
		Call Show_Error("An error occurred while calculating " & "the Kf coefficient.  This error is certainly due " & "to a value of the apparent density that is too low.")
		Error_In_Kinetic_Calculation = True
		kf = 0.00125
		Resume Exit_KF
Exit_KF: 
	End Function
	
	
	Function Re() As Double
		Dim SurfaceLoading As Double 'm/s
		Dim CarbonParticleDiameter As Double 'm
		Dim FluidViscosity As Double 'kg/m-s
		Dim FluidDensity As Double 'kg/m^3
		'	Dim Porosity As Double '(-)
		SurfaceLoading = SF()
		FluidViscosity = Bed.WaterViscosity * 0.1 'Convert g/cm-s to kg/m-s.
		FluidDensity = Bed.WaterDensity * 1000 'Convert g/cm^3 to kg/m^3.
		CarbonParticleDiameter = 2 * Carbon.ParticleRadius
		Call GetMoreBedParameters()
		Re = (CarbonParticleDiameter) * (FluidDensity) * (SurfaceLoading) / (FluidViscosity) / Bed.Porosity
	End Function
	
	
	Function SC(ByRef i As Short) As Double
		Dim VA, DA As Double
		If (Bed.Phase = 0) Then
			SC = Bed.WaterViscosity / Bed.WaterDensity / Diffl(i)
		Else
			'DA = (28.694 * Bed.Pressure) / (.08216 * (Bed.Temperature + 273.15) * 1000#)
			'VA = .0000017 * (Bed.Temperature + 273.15) ^ .818
			'Diffg in cm2/s
			'VA in g/cm.s
			'DA in g/cm3
			VA = Bed.WaterViscosity 'NOTE: WHEN BED.PHASE=1, THIS IS ACTUALLY THE _AIR_ VISCOSITY.
			DA = Bed.WaterDensity 'NOTE: WHEN BED.PHASE=1, THIS IS ACTUALLY THE _AIR_ DENSITY.
			SC = VA / (DA * Diffg(i))
		End If
	End Function
	
	
	Function SF() As Double
		Dim CrossSectionalArea As Double 'm^2
		CrossSectionalArea = 3.14159 * ((Bed.Diameter) ^ 2) / 4
		SF = Bed.Flowrate / CrossSectionalArea
	End Function
	
	
	Function Tortuosity(ByRef i As Short) As Double
		'THE OLD CORRELATION FOR Constant_Tortuosity = True
		'HAS BEEN REMOVED FOR SEVERAL YEARS NOW (1998).
		'If (Component(i).Constant_Tortuosity) Then
		'  tortuosity = 0.782 + 0.925 * CDbl(Bed.Length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#)
		'Else
		Tortuosity = 1#
		'End If
	End Function
	
	
	Function Get_Correlation_Description(ByRef Which As Short) As String
		Select Case Which
			Case 0 'FOR FILM DIFFUSION.
				Select Case Bed.Phase
					Case 0 'LIQUID-PHASE.
						Get_Correlation_Description = "Gnielinski Correlation"
					Case 1 'GAS-PHASE.
						If (USE_GASPHASE_WAKAO_AND_FUNAZUKRI) Then
							Get_Correlation_Description = "Wakao and Funazukri Correlation"
						Else
							Get_Correlation_Description = "Gnielinski Correlation"
						End If
				End Select
			Case 1 'FOR SURFACE DIFFUSION.
				Get_Correlation_Description = "Sontheimer Correlation"
			Case 2 'FOR PORE DIFFUSION.
				Select Case Bed.Phase
					Case 0 'LIQUID-PHASE.
						Get_Correlation_Description = "Hayduk and Laudie for diffusion coefficient, user-entry for tortuosity"
					Case 1 'GAS-PHASE.
						Get_Correlation_Description = "Wilke-Lee modification of the Hirschfelder - Bird - Spotz method for diffusion coefficient, user-entry for tortuosity"
				End Select
		End Select
	End Function
	
	
	
	
	
	
	
	
	
	
	Function SPDFR_Corr(ByRef i As Short) As Object
		Select Case Component(i).SPDFR_Low_Concentration
			Case True
				'UPGRADE_WARNING: Couldn't resolve default property of object SPDFR_Corr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SPDFR_Corr = 5.5797 * (Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#) ^ (-0.595)
			Case False
				'UPGRADE_WARNING: Couldn't resolve default property of object SPDFR_Corr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SPDFR_Corr = 16.263 * (Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#) ^ (-0.843)
		End Select
	End Function
	
	
	Sub Update_KP_Values()
		Dim i As Short
		For i = 1 To Number_Component
			'Update kf
			If Component(i).Corr(1) Then
				Component(i).kf = kf(i)
			End If
			'Update Ds
			If Component(i).Corr(2) Then
				Component(i).Ds = Ds(i)
			End If
			'Update Dp
			If Component(i).Corr(3) Then
				Component(i).Dp = Dp(i)
			End If
		Next i
	End Sub
End Module