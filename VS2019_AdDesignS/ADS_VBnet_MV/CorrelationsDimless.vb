Option Strict Off
Option Explicit On
Module CorrelationsDimless
	
	
	Function Bip(ByRef i As Short) As Double
		Bip = (ST(i)) / (Edp(i))
	End Function
	Function Bis(ByRef i As Short) As Double
		Bis = (ST(i)) / (Eds(i))
	End Function
	
	
	Function Dgp(ByRef i As Short) As Double
		Dim EBED, TAU As Double
		EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
		Dgp = Carbon.Porosity * (1 - EBED) / (EBED)
	End Function
	Function Dgs(ByRef i As Short) As Double
		Dim EBED As Double
		Dim qe As Double
		EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
		qe = (Component(i).Use_K) * (Component(i).InitialConcentration) ^ (Component(i).Use_OneOverN)
		Dgs = Carbon.Density * qe * (1 - EBED) * 1000# / (EBED * Component(i).InitialConcentration)
		' Explanation:
		' Dgs = (rho_a)*(q_e)*(1-epsilon)/(epsilon*C_0)
		' Note:
		' "epsilon" is EBED=1-Bed.Weight/Bed.Diameter/Bed.Diameter/Pi*4/Bed.Length/Carbon.Density/1000#
		' "C_0" is Component(i).InitialConcentration
		' "rho_a" is "adsorbent density which includes pore volume"
		' "q_e" is "adsorbent phase concentration in equilibrium
		'           with initial bulk phase concentration
		'           (K(i))*(C_0(i))^(Use_OneOverN(i))
	End Function
	
	
	Function Edp(ByRef i As Short) As Object
		Dim TAU, EBED, Dgp As Double
		EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
		TAU = Bed.Diameter * Bed.Diameter * PI / 4 * Bed.Length * EBED / Bed.Flowrate
		Dgp = Carbon.Porosity * (1 - EBED) / EBED
		'DGp = (1 - EBED) / EBED
		'UPGRADE_WARNING: Couldn't resolve default property of object Edp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Edp = Component(i).Dp * Dgp * TAU / (Carbon.ParticleRadius * 100) ^ 2
	End Function
	Function Eds(ByRef i As Short) As Object
		Dim TAU, EBED, qe, Dgs As Double
		'QE in mg/g
		'Initial Conc. in mg/l
		'Tau in s
		EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
		qe = Component(i).Use_K * (Component(i).InitialConcentration) ^ Component(i).Use_OneOverN
		TAU = Bed.Diameter * Bed.Diameter * PI / 4 * Bed.Length * EBED / Bed.Flowrate
		Dgs = Carbon.Density * 1000# * (1 - EBED) * qe / EBED / Component(i).InitialConcentration
		'UPGRADE_WARNING: Couldn't resolve default property of object Eds. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Eds = Component(i).Ds * Dgs * TAU / (Carbon.ParticleRadius * 100) ^ 2
	End Function
	
	
	Function ST(ByRef i As Short) As Double
		Dim EBED, TAU As Double
		EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
		TAU = Bed.Diameter * Bed.Diameter * PI / 4 * Bed.Length * EBED / Bed.Flowrate
		'kf in cm/s
		'Tau in s
		'PArticleRadius in m, converted in cm
		ST = Component(i).kf * (1 - EBED) * TAU / EBED / (Carbon.ParticleRadius * 100)
	End Function
End Module