Option Explicit

Sub CalculateBedArea ()
    'Column Area (m2)

     Bed.Area = GetBedArea(Bed.Diameter)

End Sub

Sub CalculateBedDensity ()

    'Bed Density (g/cm3)
    Bed.Density = GetBedDensity(Bed.Weight, Bed.Volume)

End Sub

Sub CalculateBedPorosity ()

    'Bed Porosity (-)
    Bed.Porosity = GetBedPorosity(Bed.Density, Resin.ApparentDensity)

End Sub

Sub CalculateBedVolume ()
    'Bed Volume (m3)

     Bed.Volume = GetBedVolume(Bed.Area, Bed.Length)

End Sub

Sub CalculateDimensionlessGroups ()
    'This subroutine will calculate the dimensionless groups for the ion
    'numbered NumberOfIonToEdit

    Dim i As Integer, ListIndex As Integer

    i = NumberOfIonToEdit

    If EditingCation Or AddingCation Then

       If Not OKToGetCationDimensionless Then Exit Sub

       'Calculate surface distribution parameter, Dgs
       Cation(i).Dimensionless.SurfaceDistributionParameter = GetSurfaceDistributionParameter(Resin.ApparentDensity, Resin.TotalCapacity, Bed.Porosity, SumCationInitialEquivalents)

       'Calculate pore distribution parameter, Dgp
       Cation(i).Dimensionless.PoreDistributionParameter = GetPoreDistributionParameter(Resin.ParticlePorosity, Bed.Porosity)

       'Calculate total equivalent distribution parameter, Dgt
       Cation(i).Dimensionless.TotalDistributionParameter = GetTotalDistributionParameter(Cation(i).Dimensionless.SurfaceDistributionParameter, Cation(i).Dimensionless.PoreDistributionParameter)

       'Calculate pore diffusion modulus, Edp
       Cation(i).Dimensionless.PoreDiffusionModulus = GetPoreDiffusionModulus(Cation(i).Kinetic.PoreDiffusivity.Value, Bed.EffectiveContactTime, Bed.Porosity, Resin.ParticlePorosity, Resin.ParticleRadius * 100#)

       'Calculate modified Stanton Number, St
       Cation(i).Dimensionless.StantonNumber = GetStantonNumber(Cation(i).Kinetic.IonicTransportCoefficient.Value, Bed.EffectiveContactTime, Bed.Porosity, Resin.ParticleRadius * 100#)

       'Calculate pore biot number, Bip
       Cation(i).Dimensionless.PoreBiotNumber = GetPoreBiotNumber(Cation(i).Dimensionless.PoreDiffusionModulus, Cation(i).Dimensionless.StantonNumber)

       'Generate Click Event on cboKinDimComponent
       ListIndex = frmIonExchangeMain!cboKinDimComponent.ListIndex
       frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
       frmIonExchangeMain!cboKinDimComponent.ListIndex = ListIndex

    ElseIf EditingAnion Or AddingAnion Then

       If Not OKToGetAnionDimensionless Then Exit Sub

       'Calculate surface distribution parameter, Dgs
       Anion(i).Dimensionless.SurfaceDistributionParameter = GetSurfaceDistributionParameter(Resin.ApparentDensity, Resin.TotalCapacity, Bed.Porosity, SumAnionInitialEquivalents)

       'Calculate pore distribution parameter, Dgp
       Anion(i).Dimensionless.PoreDistributionParameter = GetPoreDistributionParameter(Resin.ParticlePorosity, Bed.Porosity)

       'Calculate total equivalent distribution parameter, Dgt
       Anion(i).Dimensionless.TotalDistributionParameter = GetTotalDistributionParameter(Anion(i).Dimensionless.SurfaceDistributionParameter, Anion(i).Dimensionless.PoreDistributionParameter)

       'Calculate pore diffusion modulus, Edp
       Anion(i).Dimensionless.PoreDiffusionModulus = GetPoreDiffusionModulus(Anion(i).Kinetic.PoreDiffusivity.Value, Bed.EffectiveContactTime, Bed.Porosity, Resin.ParticlePorosity, Resin.ParticleRadius * 100#)

       'Calculate modified Stanton Number, St
       Anion(i).Dimensionless.StantonNumber = GetStantonNumber(Anion(i).Kinetic.IonicTransportCoefficient.Value, Bed.EffectiveContactTime, Bed.Porosity, Resin.ParticleRadius * 100#)

       'Calculate pore biot number, Bip
       Anion(i).Dimensionless.PoreBiotNumber = GetPoreBiotNumber(Anion(i).Dimensionless.PoreDiffusionModulus, Anion(i).Dimensionless.StantonNumber)

       'Generate Click Event on cboKinDimComponent
       ListIndex = frmIonExchangeMain!cboKinDimComponent.ListIndex
       frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
       frmIonExchangeMain!cboKinDimComponent.ListIndex = ListIndex

    End If


End Sub

Sub CalculateEffectiveContactTime ()

    Bed.EffectiveContactTime = GetBedEffectiveContactTime(Bed.Volume, Bed.Porosity, Bed.Flowrate.Value)

End Sub

Sub CalculateEquivalentInitialConc (EquivalentConc As Double, Concentration_MG_per_L As Double, Valence As Double, MolecularWeight As Double)

    EquivalentConc = Concentration_MG_per_L * ConcentrationConversionFactor(CONCENTRATION_MEQ_per_L, Valence, MolecularWeight)

End Sub

Sub CalculateInterstitialVelocity ()

    'Interstitial Velocity (cm/s)
    Bed.InterstitialVelocity = GetInterstitialVelocity(Bed.SuperficialVelocity, Bed.Porosity)

End Sub

Sub CalculateKineticParameters ()
    'This subroutine calculates kinetic parameters for one specified ion
    Dim i As Integer

    i = NumberOfIonToEdit

    If EditingCation Or AddingCation Then

       'Calculate Liquid Diffusivity
       Cation(i).Kinetic.LiquidDiffusivityCorrelation = GetLiquidDiffusivityNernstHaskell(NernstHaskell.GasConstant, Operating.Temperature, Cation(i).Kinetic.NernstHaskellCation.Valence, Cation(i).Kinetic.NernstHaskellAnion.Valence, NernstHaskell.FaradaysConstant, Cation(i).Kinetic.NernstHaskellCation.LimitingIonicConductance, Cation(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
       If Cation(i).Kinetic.LiquidDiffusivity.UserInput = False Then
          Cation(i).Kinetic.LiquidDiffusivityUserInput = Cation(i).Kinetic.LiquidDiffusivityCorrelation
          Cation(i).Kinetic.LiquidDiffusivity.Value = Cation(i).Kinetic.LiquidDiffusivityCorrelation
       End If

       'Calculate Ionic Transport Coefficient, kf
       Cation(i).Kinetic.ReynoldsNumber = GetReynoldsNumber(Resin.ParticleDiameter * 100#, Bed.InterstitialVelocity, Operating.LiquidDensity, Operating.LiquidViscosity)
       Cation(i).Kinetic.SchmidtNumber = GetSchmidtNumber(Operating.LiquidViscosity, Operating.LiquidDensity, Cation(i).Kinetic.LiquidDiffusivity.Value)
       Cation(i).Kinetic.IonicTransportCoeffCorrelation = GetIonicTransportCoefficient(Cation(i).Kinetic.LiquidDiffusivity.Value, Resin.ParticleDiameter * 100#, Bed.Porosity, Cation(i).Kinetic.ReynoldsNumber, Cation(i).Kinetic.SchmidtNumber)
       If Cation(i).Kinetic.IonicTransportCoefficient.UserInput = False Then
          Cation(i).Kinetic.IonicTransportCoeffUserInput = Cation(i).Kinetic.IonicTransportCoeffCorrelation
          Cation(i).Kinetic.IonicTransportCoefficient.Value = Cation(i).Kinetic.IonicTransportCoeffCorrelation
       End If

       'Calculate Pore Diffusivity, Dp
       Cation(i).Kinetic.PoreDiffusivityCorrelation = GetPoreDiffusivity(Cation(i).Kinetic.LiquidDiffusivity.Value, Resin.Tortuosity)
       If Cation(i).Kinetic.PoreDiffusivity.UserInput = False Then
          Cation(i).Kinetic.PoreDiffusivityUserInput = Cation(i).Kinetic.PoreDiffusivityCorrelation
          Cation(i).Kinetic.PoreDiffusivity.Value = Cation(i).Kinetic.PoreDiffusivityCorrelation
       End If

    ElseIf EditingAnion Or AddingAnion Then

       'Calculate Liquid Diffusivity, Dl
       Anion(i).Kinetic.LiquidDiffusivityCorrelation = GetLiquidDiffusivityNernstHaskell(NernstHaskell.GasConstant, Operating.Temperature, Anion(i).Kinetic.NernstHaskellCation.Valence, Anion(i).Kinetic.NernstHaskellAnion.Valence, NernstHaskell.FaradaysConstant, Anion(i).Kinetic.NernstHaskellCation.LimitingIonicConductance, Anion(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
       If Anion(i).Kinetic.LiquidDiffusivity.UserInput = False Then
          Anion(i).Kinetic.LiquidDiffusivityUserInput = Anion(i).Kinetic.LiquidDiffusivityCorrelation
          Anion(i).Kinetic.LiquidDiffusivity.Value = Anion(i).Kinetic.LiquidDiffusivityCorrelation
       End If

       'Calculate Ionic Transport Coefficient, kf
       Anion(i).Kinetic.ReynoldsNumber = GetReynoldsNumber(Resin.ParticleDiameter, Bed.InterstitialVelocity, Operating.LiquidDensity, Operating.LiquidViscosity)
       Anion(i).Kinetic.SchmidtNumber = GetSchmidtNumber(Operating.LiquidViscosity, Operating.LiquidDensity, Anion(i).Kinetic.LiquidDiffusivity.Value)
       Anion(i).Kinetic.IonicTransportCoeffCorrelation = GetIonicTransportCoefficient(Anion(i).Kinetic.LiquidDiffusivity.Value, Resin.ParticleDiameter * 100#, Bed.Porosity, Anion(i).Kinetic.ReynoldsNumber, Anion(i).Kinetic.SchmidtNumber)
       If Anion(i).Kinetic.IonicTransportCoefficient.UserInput = False Then
          Anion(i).Kinetic.IonicTransportCoeffUserInput = Anion(i).Kinetic.IonicTransportCoeffCorrelation
          Anion(i).Kinetic.IonicTransportCoefficient.Value = Anion(i).Kinetic.IonicTransportCoeffCorrelation
       End If

       'Calculate Pore Diffusivity, Dp
       Anion(i).Kinetic.PoreDiffusivityCorrelation = GetPoreDiffusivity(Anion(i).Kinetic.LiquidDiffusivity.Value, Resin.Tortuosity)
       If Anion(i).Kinetic.PoreDiffusivity.UserInput = False Then
          Anion(i).Kinetic.PoreDiffusivityUserInput = Anion(i).Kinetic.PoreDiffusivityCorrelation
          Anion(i).Kinetic.PoreDiffusivity.Value = Anion(i).Kinetic.PoreDiffusivityCorrelation
       End If

    End If

End Sub

Sub CalculateLiquidDensity ()
    'Returns Operating.LiquidDensity in g/cm3

    Dim LiquidDensity As Double

    Call H2ODens(LiquidDensity, Operating.Temperature)

    'Convert LiquidDensity from kg/m3 to g/cm3
    LiquidDensity = LiquidDensity / 1000#

    Operating.LiquidDensity = LiquidDensity

End Sub

Sub CalculateLiquidViscosity ()
    'Returns Operating.LiquidViscosity in g/cm3

    Dim LiquidViscosity As Double

    Call H2OVisc(LiquidViscosity, Operating.Temperature)

    'Convert LiquidViscosity from kg/m/sec to g/cm/sec
    LiquidViscosity = LiquidViscosity * 1000# / 100#

    Operating.LiquidViscosity = LiquidViscosity

End Sub

Sub CalculateParticleDiameter ()

    Resin.ParticleDiameter = GetParticleDiameter(Resin.ParticleRadius)

End Sub

Sub CalculateSeparationFactors ()
    Dim i As Integer, j As Integer

       For i = 1 To NumberOfIons
           For j = 1 To NumberOfIons
               If SeparationFactorInput.Row = True Then
                  TwoDimSeparationFactors(i, j) = OneDimSeparationFactors(i) / OneDimSeparationFactors(j)
               Else
                  TwoDimSeparationFactors(i, j) = OneDimSeparationFactors(j) / OneDimSeparationFactors(i)
               End If
           Next j
       Next i

End Sub

Sub CalculateSumEquivInitialConc ()
    'This subroutine will calculate the sum of the time-averaged initial influent concentrations
    Dim i As Integer, NumSelected As Integer

    If EditingCation Or AddingCation Then
       SumCationInitialEquivalents = 0
       For i = 1 To NumSelectedCations
           NumSelected = Cations_Selected(i)
           SumCationInitialEquivalents = SumCationInitialEquivalents + Cation(NumSelected).EquivalentInitialConcentration
       Next i
       If HaveValue(SumCationInitialEquivalents) Then
          OKToGetCationDimensionless = True
       Else
          OKToGetCationDimensionless = False
       End If
    ElseIf EditingAnion Or AddingAnion Then
       SumAnionInitialEquivalents = 0
       For i = 1 To NumSelectedAnions
           NumSelected = Anions_Selected(i)
           SumAnionInitialEquivalents = SumAnionInitialEquivalents + Anion(NumSelected).EquivalentInitialConcentration
       Next i
       If HaveValue(SumAnionInitialEquivalents) Then
          OKToGetAnionDimensionless = True
       Else
          OKToGetAnionDimensionless = False
       End If
    End If

End Sub

Sub CalculateSuperficialVelocity ()

    'Superficial Velocity (cm/s)
    Bed.SuperficialVelocity = GetSuperficialVelocity(Bed.Flowrate.Value, Bed.Area)

End Sub

Function GetBedArea (BedDiameter As Double) As Double
   'Input:
   '   BedDiameter in m
   'Output
   '   BedArea in m2

   GetBedArea = Pi * BedDiameter ^ 2 / 4

End Function

Function GetBedDensity (BedWeight As Double, BedVolume As Double) As Double
   'Input:
   '   BedWeight in kg
   '   BedVolume in m3
   'Output:
   '   Bed Density in g/cm3

   GetBedDensity = (BedWeight / BedVolume) / 1000

End Function

Function GetBedEffectiveContactTime (BedVolume As Double, BedPorosity As Double, BedFlowrate As Double) As Double
   'Input:
   '   BedVolume in m3
   '   BedPorosity in (-)
   '   BedFlowrate in m3/s
   'Output:
   '   BedEffectiveContactTime in s

   GetBedEffectiveContactTime = BedVolume * BedPorosity / BedFlowrate

End Function

Function GetBedPorosity (BedDensity As Double, ResinApparentDensity As Double)
   'Input:
   '   BedDensity in g/cm3
   '   ResinApparentDensity in g/cm3
   'Output:
   '   BedPorosity in (-)

   GetBedPorosity = 1# - (BedDensity / ResinApparentDensity)
       
End Function

Function GetBedVolume (BedArea As Double, BedLength As Double)
   'Input:
   '   BedArea in m2
   '   BedLength in m
   'Output
   '   BedVolume in m3

   GetBedVolume = BedArea * BedLength

End Function

Function GetInterstitialVelocity (SuperficialVelocity As Double, BedPorosity As Double) As Double
   'Input:
   '   SuperficialVelocity in cm/s
   '   BedPorosity in (-)
   'Output:
   '   InterstitialVelocity in cm/s

   GetInterstitialVelocity = SuperficialVelocity / BedPorosity

End Function

Function GetIonicTransportCoefficient (DiffusionCoefficient As Double, ParticleDiameter As Double, BedPorosity As Double, ReynoldsNumber As Double, SchmidtNumber As Double) As Double
    'Input:
    '   DiffusionCoefficient in cm2/s
    '   ParticleDiameter in cm
    '   Bed Porosity as (-)
    '   ReynoldsNumber as (-)
    '   SchmidtNumber as (-)
    'Output:
    '   IonicTransportCoefficient in cm/s

    If IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1 Then   'Wildhagen Correlation
       GetIonicTransportCoefficient = (DiffusionCoefficient / ParticleDiameter) * (.86 / BedPorosity) * (ReynoldsNumber ^ (1 / 2)) * (SchmidtNumber ^ (1 / 3))
    ElseIf IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2 Then   'Gnielinski Correlation
       GetIonicTransportCoefficient = ((1# + 1.5 * (1# - BedPorosity)) * DiffusionCoefficient / ParticleDiameter) * (2# + .644 * (ReynoldsNumber ^ (1 / 2)) * (SchmidtNumber ^ (1 / 3)))
    End If

End Function

Function GetLiquidDiffusivityNernstHaskell (GasConstant As Double, Temperature As Double, CationValence As Double, AnionValence As Double, FaradaysConstant As Double, CationLimitingIonicConductance As Double, AnionLimitingIonicConductance As Double) As Double
   'Input:
   '   Gas Constant in J/mol/K
   '   Temperature in K
   '   CationValence in eq/mol
   '   AnionValence in eq/mol
   '   FaradaysConstant in cal/g/eq
   '   CationLimitingIonicConductance in (A/cm2) (V/cm) (g-eq/cm3)
   '   AnionLimitingIonicConductance in (A/cm2) (V/cm) (g-eq/cm3)
   'Output:
   '   DiffusionCoefficient in cm2/s

   GetLiquidDiffusivityNernstHaskell = (GasConstant * Temperature * ((1 / CationValence) + (1 / AnionValence))) / FaradaysConstant ^ 2 / ((1 / CationLimitingIonicConductance) + (1 / AnionLimitingIonicConductance))

End Function

Function GetParticleDiameter (ParticleRadius As Double) As Double
   'Input:
   '   ParticleRadius in m
   'Output:
   '   ParticleDiameter in m

   GetParticleDiameter = 2# * ParticleRadius

End Function

Function GetPoreBiotNumber (PoreDiffusionModulus As Double, StantonNumber As Double) As Double
   'Input:
   '   PoreDiffusionModulus in (-)
   '   SurfaceDiffusionModulus in (-)
   'Output:
   '   PoreBiotNumber in (-)

    GetPoreBiotNumber = StantonNumber / PoreDiffusionModulus

End Function

Function GetPoreDiffusionModulus (PoreDiffusivity As Double, BedEffectiveContactTime As Double, BedPorosity As Double, ParticlePorosity As Double, ParticleRadius As Double) As Double
   'Input:
   '   PoreDiffusivity in cm2/s
   '   BedEffectiveContactTime in s
   '   BedPorosity in (-)
   '   ParticlePorosity in (-)
   '   ParticleRadius in cm
   'Output:
   '   PoreDiffusionModulus in (-)

     GetPoreDiffusionModulus = PoreDiffusivity * BedEffectiveContactTime * (1# - BedPorosity) * ParticlePorosity / (ParticleRadius ^ 2) / BedPorosity

End Function

Function GetPoreDiffusivity (DiffusionCoefficient As Double, ResinTortuosity As Double) As Double
    'Input:
    '   DiffusiionCoefficient in cm2/s
    '   ResinTortuosity in Dimensionless
    'Output:
    '   PoreDiffusivity in cm2/s

    GetPoreDiffusivity = DiffusionCoefficient / ResinTortuosity

End Function

Function GetPoreDistributionParameter (ParticlePorosity As Double, BedPorosity As Double) As Double
   'Input:
   '   ParticlePorosity as (-)
   '   BedPorosity as (-)
   'Output:
   '   PoreDistributionParameter as (-)

    GetPoreDistributionParameter = ParticlePorosity * (1# - BedPorosity) / BedPorosity

End Function

Function GetReynoldsNumber (ParticleDiameter As Double, InterstitialVelocity As Double, LiquidDensity As Double, LiquidViscosity As Double) As Double
    'Input:
    '   ParticleDiameter in cm
    '   InterstitialVelocity in cm/s
    '   LiquidDensity in g/cm3
    '   LiquidViscosity in g/cm/s
    'Output:
    '   ReynoldsNumber in (-)

    GetReynoldsNumber = ParticleDiameter * InterstitialVelocity * LiquidDensity / LiquidViscosity

End Function

Function GetSchmidtNumber (LiquidViscosity As Double, LiquidDensity As Double, DiffusionCoefficient As Double) As Double
    'Input:
    '   LiquidViscosity in g/cm/s
    '   LiquidDensity in g/cm3
    '   DiffusionCoefficient in cm2/s
    'Output:
    '   SchmidtNumber in (-)

    GetSchmidtNumber = LiquidViscosity / LiquidDensity / DiffusionCoefficient

End Function

Function GetStantonNumber (IonicTransportCoefficient As Double, BedEffectiveContactTime As Double, BedPorosity As Double, ParticleRadius As Double) As Double
   'Input:
   '   IonicTransportCoefficient in cm/s
   '   BedEffectiveContactTime in s
   '   BedPorosity in (-)
   '   ParticleRadius in cm
   'Output:
   '   StantonNumber in (-)

    GetStantonNumber = IonicTransportCoefficient * BedEffectiveContactTime * (1# - BedPorosity) / ParticleRadius / BedPorosity

End Function

Function GetSuperficialVelocity (InletFlowrate As Double, BedArea As Double) As Double
   'Input:
   '   InletFlowrate in m3/s
   '   BedArea in m2
   'Output:
   '   SuperficialVelocity in cm/s

   GetSuperficialVelocity = (InletFlowrate / BedArea) * 100

End Function

Function GetSurfaceDistributionParameter (ApparentDensity As Double, TotalResinCapacity, BedPorosity As Double, SumOfInitialEquivalentConcs) As Double
   'Input:
   '   ApparentDensity in g/cm3
   '   TotalResinCapacity in meq/g
   '   BedPorosity in (-)
   '   SumOfInitialEquivalentConcs in meq/L
   'Output:
   '   SurfaceDistributionParameter in (-)

    GetSurfaceDistributionParameter = (ApparentDensity * 1000#) * TotalResinCapacity * (1# - BedPorosity) / BedPorosity / SumOfInitialEquivalentConcs

End Function

Function GetTotalDistributionParameter (SurfaceDistributionParameter As Double, PoreDistributionParameter As Double) As Double
   'Input:
   '   SurfaceDistributionParameter as (-)
   '   PoreDistributionParameter as (-)
   'Output:
   '   TotalDistributionParameter as (-)

    GetTotalDistributionParameter = SurfaceDistributionParameter + PoreDistributionParameter

End Function

Sub UpdateDimensionlessGroupAllIons ()
    Dim i As Integer

    EditingAnion = False
    EditingCation = True
    For i = 1 To NumSelectedCations
        NumberOfIonToEdit = Cations_Selected(i)
        Call CalculateDimensionlessGroups
    Next i
    EditingCation = False
    EditingAnion = True
    For i = 1 To NumSelectedAnions
        NumberOfIonToEdit = Anions_Selected(i)
        Call CalculateDimensionlessGroups
    Next i
    EditingAnion = False

End Sub

Sub UpdateKineticParametersAllIons ()
    Dim i As Integer

    EditingAnion = False
    EditingCation = True
    For i = 1 To NumberOfCations
        NumberOfIonToEdit = i
        Call CalculateKineticParameters
    Next i
    EditingCation = False
    EditingAnion = True
    For i = 1 To NumberOfAnions
        NumberOfIonToEdit = i
        Call CalculateKineticParameters
    Next i
    EditingAnion = False

End Sub

