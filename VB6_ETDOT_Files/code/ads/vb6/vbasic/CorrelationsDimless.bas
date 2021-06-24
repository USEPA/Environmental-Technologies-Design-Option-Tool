Attribute VB_Name = "CorrelationsDimless"
Option Explicit


Function Bip(i As Integer) As Double
  Bip = (ST(i)) / (Edp(i))
End Function
Function Bis(i As Integer) As Double
  Bis = (ST(i)) / (Eds(i))
End Function


Function Dgp(i As Integer) As Double
Dim EBED As Double, TAU As Double
  EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
  Dgp = Carbon.Porosity * (1 - EBED) / (EBED)
End Function
Function Dgs(i As Integer) As Double
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


Function Edp(i As Integer)
Dim EBED As Double, TAU As Double, Dgp As Double
  EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
  TAU = Bed.Diameter * Bed.Diameter * PI / 4 * Bed.Length * EBED / Bed.Flowrate
  Dgp = Carbon.Porosity * (1 - EBED) / EBED
  'DGp = (1 - EBED) / EBED
  Edp = Component(i).Dp * Dgp * TAU / (Carbon.ParticleRadius * 100) ^ 2
End Function
Function Eds(i As Integer)
Dim EBED As Double, qe As Double, TAU As Double, Dgs As Double
  'QE in mg/g
  'Initial Conc. in mg/l
  'Tau in s
  EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
  qe = Component(i).Use_K * (Component(i).InitialConcentration) ^ Component(i).Use_OneOverN
  TAU = Bed.Diameter * Bed.Diameter * PI / 4 * Bed.Length * EBED / Bed.Flowrate
  Dgs = Carbon.Density * 1000# * (1 - EBED) * qe / EBED / Component(i).InitialConcentration
  Eds = Component(i).Ds * Dgs * TAU / (Carbon.ParticleRadius * 100) ^ 2
End Function


Function ST(i As Integer) As Double
Dim EBED As Double, TAU As Double
  EBED = 1 - Bed.Weight / Bed.Diameter / Bed.Diameter / PI * 4 / Bed.Length / Carbon.Density / 1000#
  TAU = Bed.Diameter * Bed.Diameter * PI / 4 * Bed.Length * EBED / Bed.Flowrate
  'kf in cm/s
  'Tau in s
  'PArticleRadius in m, converted in cm
  ST = Component(i).kf * (1 - EBED) * TAU / EBED / (Carbon.ParticleRadius * 100)
End Function

