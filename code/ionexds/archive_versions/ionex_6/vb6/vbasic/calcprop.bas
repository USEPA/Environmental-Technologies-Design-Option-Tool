Attribute VB_Name = "calcprop"
Option Explicit

Sub CalculateBedArea()
    'Column Area (m2)


     NowProj.Bed.Area = GetBedArea(NowProj.Bed.Diameter)

End Sub

Sub CalculateBedDensity()

    'Bed Density (g/cm3)
    NowProj.Bed.Density = GetBedDensity(NowProj.Bed.Weight, NowProj.Bed.Volume)

End Sub

Sub CalculateBedPorosity()

    'Bed Porosity (-)
    NowProj.Bed.Porosity = GetBedPorosity(NowProj.Bed.Density, NowProj.Resin.ApparentDensity)

End Sub

Sub CalculateBedVolume()
    'Bed Volume (m3)

     NowProj.Bed.Volume = GetBedVolume(NowProj.Bed.Area, NowProj.Bed.length)

End Sub

Sub CalculateDimensionlessGroups()
    'This subroutine will calculate the dimensionless groups for the ion
    'numbered NumberOfIonToEdit

    Dim i As Integer, ListIndex As Integer

    i = NumberOfIonToEdit

    If EditingCation Or AddingCation Then

       If Not NowProj.OKToGetCationDimensionless Then Exit Sub

       'Calculate surface distribution parameter, Dgs
       NowProj.Cation(i).Dimensionless.SurfaceDistributionParameter = GetSurfaceDistributionParameter(NowProj.Resin.ApparentDensity, NowProj.Resin.TotalCapacity, NowProj.Bed.Porosity, NowProj.SumCationInitialEquivalents)

       'Calculate pore distribution parameter, Dgp
       NowProj.Cation(i).Dimensionless.PoreDistributionParameter = GetPoreDistributionParameter(NowProj.Resin.ParticlePorosity, NowProj.Bed.Porosity)

       'Calculate total equivalent distribution parameter, Dgt
       NowProj.Cation(i).Dimensionless.TotalDistributionParameter = GetTotalDistributionParameter(NowProj.Cation(i).Dimensionless.SurfaceDistributionParameter, NowProj.Cation(i).Dimensionless.PoreDistributionParameter)

       'Calculate pore diffusion modulus, Edp
       NowProj.Cation(i).Dimensionless.PoreDiffusionModulus = GetPoreDiffusionModulus(NowProj.Cation(i).Kinetic.PoreDiffusivity.Value, NowProj.Bed.EffectiveContactTime, NowProj.Bed.Porosity, NowProj.Resin.ParticlePorosity, NowProj.Resin.ParticleRadius * 100#)

       'Calculate modified Stanton Number, St
       NowProj.Cation(i).Dimensionless.StantonNumber = GetStantonNumber(NowProj.Cation(i).Kinetic.IonicTransportCoefficient.Value, NowProj.Bed.EffectiveContactTime, NowProj.Bed.Porosity, NowProj.Resin.ParticleRadius * 100#)

       'Calculate pore biot number, Bip
       NowProj.Cation(i).Dimensionless.PoreBiotNumber = GetPoreBiotNumber(NowProj.Cation(i).Dimensionless.PoreDiffusionModulus, NowProj.Cation(i).Dimensionless.StantonNumber)

       'Generate Click Event on cboKinDimComponent
       ListIndex = frmIonExchangeMain!cboKinDimComponent.ListIndex
'       frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
       frmIonExchangeMain!cboKinDimComponent.ListIndex = ListIndex

    ElseIf EditingAnion Or AddingAnion Then

       If Not NowProj.OKToGetAnionDimensionless Then Exit Sub

       'Calculate surface distribution parameter, Dgs
       NowProj.Anion(i).Dimensionless.SurfaceDistributionParameter = GetSurfaceDistributionParameter(NowProj.Resin.ApparentDensity, NowProj.Resin.TotalCapacity, NowProj.Bed.Porosity, NowProj.SumAnionInitialEquivalents)

       'Calculate pore distribution parameter, Dgp
       NowProj.Anion(i).Dimensionless.PoreDistributionParameter = GetPoreDistributionParameter(NowProj.Resin.ParticlePorosity, NowProj.Bed.Porosity)

       'Calculate total equivalent distribution parameter, Dgt
       NowProj.Anion(i).Dimensionless.TotalDistributionParameter = GetTotalDistributionParameter(NowProj.Anion(i).Dimensionless.SurfaceDistributionParameter, NowProj.Anion(i).Dimensionless.PoreDistributionParameter)

       'Calculate pore diffusion modulus, Edp
       NowProj.Anion(i).Dimensionless.PoreDiffusionModulus = GetPoreDiffusionModulus(NowProj.Anion(i).Kinetic.PoreDiffusivity.Value, NowProj.Bed.EffectiveContactTime, NowProj.Bed.Porosity, NowProj.Resin.ParticlePorosity, NowProj.Resin.ParticleRadius * 100#)

       'Calculate modified Stanton Number, St
       NowProj.Anion(i).Dimensionless.StantonNumber = GetStantonNumber(NowProj.Anion(i).Kinetic.IonicTransportCoefficient.Value, NowProj.Bed.EffectiveContactTime, NowProj.Bed.Porosity, NowProj.Resin.ParticleRadius * 100#)

       'Calculate pore biot number, Bip
       NowProj.Anion(i).Dimensionless.PoreBiotNumber = GetPoreBiotNumber(NowProj.Anion(i).Dimensionless.PoreDiffusionModulus, NowProj.Anion(i).Dimensionless.StantonNumber)

       'Generate Click Event on cboKinDimComponent
       ListIndex = frmIonExchangeMain!cboKinDimComponent.ListIndex
'       frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
       frmIonExchangeMain!cboKinDimComponent.ListIndex = ListIndex

    End If


End Sub

Sub CalculateEffectiveContactTime()

    NowProj.Bed.EffectiveContactTime = GetBedEffectiveContactTime(NowProj.Bed.Volume, NowProj.Bed.Porosity, NowProj.Bed.Flowrate.Value)

End Sub

Sub CalculateEquivalentInitialConc(EquivalentConc As Double, CONCENTRATION_MG_per_L As Double, Valence As Double, MolecularWeight As Double)

    EquivalentConc = CONCENTRATION_MG_per_L * ConcentrationConversionFactor(CONCENTRATION_MEQ_per_L, Valence, MolecularWeight)

End Sub

Sub CalculateInterstitialVelocity()

    'Interstitial Velocity (cm/s)
    NowProj.Bed.InterstitialVelocity = GetInterstitialVelocity(NowProj.Bed.SuperficialVelocity, NowProj.Bed.Porosity)

End Sub

Sub CalculateKineticParameters()
    'This subroutine calculates kinetic parameters for one specified ion
    Dim i As Integer

    i = NumberOfIonToEdit

    If EditingCation Or AddingCation Then

       'Calculate Liquid Diffusivity
       NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation = _
            GetLiquidDiffusivityNernstHaskell(NernstHaskell.GasConstant, _
            NowProj.Operating.Temperature, _
            NowProj.Cation(i).Kinetic.NernstHaskellCation.Valence, _
            NowProj.Cation(i).Kinetic.NernstHaskellAnion.Valence, _
            NernstHaskell.FaradaysConstant, _
            NowProj.Cation(i).Kinetic.NernstHaskellCation.LimitingIonicConductance, _
            NowProj.Cation(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
       If NowProj.Cation(i).Kinetic.LiquidDiffusivity.UserInput = False Then
          NowProj.Cation(i).Kinetic.LiquidDiffusivityUserInput = _
            NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation
          NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value = _
            NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation
       End If

       'Calculate Ionic Transport Coefficient, kf
       NowProj.Cation(i).Kinetic.ReynoldsNumber = _
           GetReynoldsNumber(NowProj.Resin.ParticleDiameter * 100#, _
           NowProj.Bed.InterstitialVelocity, _
           NowProj.Operating.LiquidDensity, _
           NowProj.Operating.LiquidViscosity)
       NowProj.Cation(i).Kinetic.SchmidtNumber = _
           GetSchmidtNumber(NowProj.Operating.LiquidViscosity, _
           NowProj.Operating.LiquidDensity, _
           NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value)
       NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation = _
           GetIonicTransportCoefficient( _
           NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value, _
           NowProj.Resin.ParticleDiameter * 100#, _
           NowProj.Bed.Porosity, _
           NowProj.Cation(i).Kinetic.ReynoldsNumber, _
           NowProj.Cation(i).Kinetic.SchmidtNumber)
       If NowProj.Cation(i).Kinetic.IonicTransportCoefficient.UserInput = False Then
          NowProj.Cation(i).Kinetic.IonicTransportCoeffUserInput = _
            NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation
          NowProj.Cation(i).Kinetic.IonicTransportCoefficient.Value = _
            NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation
       End If

       'Calculate Pore Diffusivity, Dp
       NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation = _
            GetPoreDiffusivity(NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value, _
            NowProj.Resin.Tortuosity)
       If NowProj.Cation(i).Kinetic.PoreDiffusivity.UserInput = False Then
          NowProj.Cation(i).Kinetic.PoreDiffusivityUserInput = _
                NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation
          NowProj.Cation(i).Kinetic.PoreDiffusivity.Value = _
                NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation
       End If

    ElseIf EditingAnion Or AddingAnion Then

       'Calculate Liquid Diffusivity, Dl
       NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation = _
            GetLiquidDiffusivityNernstHaskell(NernstHaskell.GasConstant, _
            NowProj.Operating.Temperature, _
            NowProj.Anion(i).Kinetic.NernstHaskellCation.Valence, _
            NowProj.Anion(i).Kinetic.NernstHaskellAnion.Valence, _
            NernstHaskell.FaradaysConstant, _
            NowProj.Anion(i).Kinetic.NernstHaskellCation.LimitingIonicConductance, _
            NowProj.Anion(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
       If NowProj.Anion(i).Kinetic.LiquidDiffusivity.UserInput = False Then
          NowProj.Anion(i).Kinetic.LiquidDiffusivityUserInput = _
            NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation
          NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value = _
            NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation
       End If

       'Calculate Ionic Transport Coefficient, kf
'------BEGIN MODIFICATION HOKANSON 10-AUG-2000: PFPDM_VBASIC3_v20000810
''''       nowproj.anion(i).Kinetic.ReynoldsNumber = GetReynoldsNumber(nowproj.resin.ParticleDiameter, nowproj.bed.InterstitialVelocity,  nowproj.Operating.LiquidDensity,  nowproj.Operating.LiquidViscosity)
       NowProj.Anion(i).Kinetic.ReynoldsNumber = GetReynoldsNumber(NowProj.Resin.ParticleDiameter * 100#, NowProj.Bed.InterstitialVelocity, NowProj.Operating.LiquidDensity, NowProj.Operating.LiquidViscosity)
'------END MODIFICATION HOKANSON 10-AUG-2000: PFPDM_VBASIC3_v20000810
       NowProj.Anion(i).Kinetic.SchmidtNumber = GetSchmidtNumber(NowProj.Operating.LiquidViscosity, NowProj.Operating.LiquidDensity, NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value)
       NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation = GetIonicTransportCoefficient(NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value, NowProj.Resin.ParticleDiameter * 100#, NowProj.Bed.Porosity, NowProj.Anion(i).Kinetic.ReynoldsNumber, NowProj.Anion(i).Kinetic.SchmidtNumber)
       If NowProj.Anion(i).Kinetic.IonicTransportCoefficient.UserInput = False Then
          NowProj.Anion(i).Kinetic.IonicTransportCoeffUserInput = NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation
          NowProj.Anion(i).Kinetic.IonicTransportCoefficient.Value = NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation
       End If

       'Calculate Pore Diffusivity, Dp
       NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation = GetPoreDiffusivity(NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value, NowProj.Resin.Tortuosity)
       If NowProj.Anion(i).Kinetic.PoreDiffusivity.UserInput = False Then
          NowProj.Anion(i).Kinetic.PoreDiffusivityUserInput = NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation
          NowProj.Anion(i).Kinetic.PoreDiffusivity.Value = NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation
       End If

    End If

End Sub

Sub CalculateLiquidDensity()
    'Returns  nowproj.Operating.LiquidDensity in g/cm3

    Dim LiquidDensity As Double

    '*'*'*'*
    'Dave H. is going to rewrite this dll, uncomment this code when completed
'    Call H2ODens(LiquidDensity,  nowproj.Operating.Temperature)
  LiquidDensity = 1#
    '*'*'*'*
    
    'Convert LiquidDensity from kg/m3 to g/cm3
    LiquidDensity = LiquidDensity / 1000#

     NowProj.Operating.LiquidDensity = LiquidDensity

End Sub

Sub CalculateLiquidViscosity()
    'Returns  nowproj.Operating.LiquidViscosity in g/cm3

    Dim LiquidViscosity As Double

  '*'*'*'*
  'Dave H. is going to rewrite this dll, uncomment this code when completed
'    Call H2OVisc(LiquidViscosity,  nowproj.Operating.Temperature)
  LiquidViscosity = 0.01
  '*'*'*'*
  
    'Convert LiquidViscosity from kg/m/sec to g/cm/sec
    LiquidViscosity = LiquidViscosity * 1000# / 100#

     NowProj.Operating.LiquidViscosity = LiquidViscosity

End Sub

Sub CalculateParticleDiameter()

    NowProj.Resin.ParticleDiameter = GetParticleDiameter(NowProj.Resin.ParticleRadius)

End Sub

Sub CalculateSeparationFactors()
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

Sub CalculateSumEquivInitialConc()
    'This subroutine will calculate the sum of the time-averaged initial influent concentrations
    Dim i As Integer, NumSelected As Integer

    If EditingCation Or AddingCation Then
       NowProj.SumCationInitialEquivalents = 0
       For i = 1 To NumSelectedCations
           NumSelected = Cations_Selected(i)
           NowProj.SumCationInitialEquivalents = NowProj.SumCationInitialEquivalents + NowProj.Cation(NumSelected).EquivalentInitialConcentration
       Next i
       If HaveValue(NowProj.SumCationInitialEquivalents) Then
           NowProj.OKToGetCationDimensionless = True
       Else
           NowProj.OKToGetCationDimensionless = False
       End If
    ElseIf EditingAnion Or AddingAnion Then
       NowProj.SumAnionInitialEquivalents = 0
       For i = 1 To NumSelectedAnions
           NumSelected = Anions_Selected(i)
           NowProj.SumAnionInitialEquivalents = NowProj.SumAnionInitialEquivalents + NowProj.Anion(NumSelected).EquivalentInitialConcentration
       Next i
       If HaveValue(NowProj.SumAnionInitialEquivalents) Then
           NowProj.OKToGetAnionDimensionless = True
       Else
           NowProj.OKToGetAnionDimensionless = False
       End If
    End If

End Sub

Sub CalculateSuperficialVelocity()

    'Superficial Velocity (cm/s)
    NowProj.Bed.SuperficialVelocity = GetSuperficialVelocity(NowProj.Bed.Flowrate.Value, NowProj.Bed.Area)

End Sub

Function GetBedArea(BedDiameter As Double) As Double
   'Input:
   '   BedDiameter in m
   'Output
   '   BedArea in m2

   GetBedArea = Pi * BedDiameter ^ 2 / 4

End Function

Function GetBedDensity(BedWeight As Double, BedVolume As Double) As Double
   'Input:
   '   BedWeight in kg
   '   BedVolume in m3
   'Output:
   '   Bed Density in g/cm3

   GetBedDensity = (BedWeight / BedVolume) / 1000

End Function

Function GetBedEffectiveContactTime(BedVolume As Double, BedPorosity As Double, BedFlowrate As Double) As Double
   'Input:
   '   BedVolume in m3
   '   BedPorosity in (-)
   '   BedFlowrate in m3/s
   'Output:
   '   BedEffectiveContactTime in s

   GetBedEffectiveContactTime = BedVolume * BedPorosity / BedFlowrate

End Function

Function GetBedPorosity(BedDensity As Double, ResinApparentDensity As Double)
   'Input:
   '   BedDensity in g/cm3
   '   ResinApparentDensity in g/cm3
   'Output:
   '   BedPorosity in (-)

   GetBedPorosity = 1# - (BedDensity / ResinApparentDensity)
       
End Function

Function GetBedVolume(BedArea As Double, BedLength As Double)
   'Input:
   '   BedArea in m2
   '   BedLength in m
   'Output
   '   BedVolume in m3

   GetBedVolume = BedArea * BedLength

End Function

Function GetInterstitialVelocity(SuperficialVelocity As Double, BedPorosity As Double) As Double
   'Input:
   '   SuperficialVelocity in cm/s
   '   BedPorosity in (-)
   'Output:
   '   InterstitialVelocity in cm/s

   GetInterstitialVelocity = SuperficialVelocity / BedPorosity

End Function

Function GetIonicTransportCoefficient(DiffusionCoefficient As Double, ParticleDiameter As Double, BedPorosity As Double, ReynoldsNumber As Double, SchmidtNumber As Double) As Double
    'Input:
    '   DiffusionCoefficient in cm2/s
    '   ParticleDiameter in cm
    '   Bed Porosity as (-)
    '   ReynoldsNumber as (-)
    '   SchmidtNumber as (-)
    'Output:
    '   IonicTransportCoefficient in cm/s

    If NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1 Then   'Wildhagen Correlation
       GetIonicTransportCoefficient = (DiffusionCoefficient / ParticleDiameter) * (0.86 / BedPorosity) * (ReynoldsNumber ^ (1 / 2)) * (SchmidtNumber ^ (1 / 3))
    ElseIf NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2 Then   'Gnielinski Correlation
       GetIonicTransportCoefficient = ((1# + 1.5 * (1# - BedPorosity)) * DiffusionCoefficient / ParticleDiameter) * (2# + 0.644 * (ReynoldsNumber ^ (1 / 2)) * (SchmidtNumber ^ (1 / 3)))
    End If

End Function

Function GetLiquidDiffusivityNernstHaskell(GasConstant As Double, Temperature As Double, CationValence As Double, AnionValence As Double, FaradaysConstant As Double, CationLimitingIonicConductance As Double, AnionLimitingIonicConductance As Double) As Double
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

Function GetParticleDiameter(ParticleRadius As Double) As Double
   'Input:
   '   ParticleRadius in m
   'Output:
   '   ParticleDiameter in m

   GetParticleDiameter = 2# * ParticleRadius

End Function

Function GetPoreBiotNumber(PoreDiffusionModulus As Double, StantonNumber As Double) As Double
   'Input:
   '   PoreDiffusionModulus in (-)
   '   SurfaceDiffusionModulus in (-)
   'Output:
   '   PoreBiotNumber in (-)

    GetPoreBiotNumber = StantonNumber / PoreDiffusionModulus

End Function

Function GetPoreDiffusionModulus(PoreDiffusivity As Double, BedEffectiveContactTime As Double, BedPorosity As Double, ParticlePorosity As Double, ParticleRadius As Double) As Double
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

Function GetPoreDiffusivity(DiffusionCoefficient As Double, ResinTortuosity As Double) As Double
    'Input:
    '   DiffusiionCoefficient in cm2/s
    '   ResinTortuosity in Dimensionless
    'Output:
    '   PoreDiffusivity in cm2/s

    GetPoreDiffusivity = DiffusionCoefficient / ResinTortuosity

End Function

Function GetPoreDistributionParameter(ParticlePorosity As Double, BedPorosity As Double) As Double
   'Input:
   '   ParticlePorosity as (-)
   '   BedPorosity as (-)
   'Output:
   '   PoreDistributionParameter as (-)

    GetPoreDistributionParameter = ParticlePorosity * (1# - BedPorosity) / BedPorosity

End Function

Function GetReynoldsNumber(ParticleDiameter As Double, InterstitialVelocity As Double, LiquidDensity As Double, LiquidViscosity As Double) As Double
    'Input:
    '   ParticleDiameter in cm
    '   InterstitialVelocity in cm/s
    '   LiquidDensity in g/cm3
    '   LiquidViscosity in g/cm/s
    'Output:
    '   ReynoldsNumber in (-)

    GetReynoldsNumber = ParticleDiameter * InterstitialVelocity * LiquidDensity / LiquidViscosity

End Function

Function GetSchmidtNumber(LiquidViscosity As Double, LiquidDensity As Double, DiffusionCoefficient As Double) As Double
    'Input:
    '   LiquidViscosity in g/cm/s
    '   LiquidDensity in g/cm3
    '   DiffusionCoefficient in cm2/s
    'Output:
    '   SchmidtNumber in (-)

    GetSchmidtNumber = LiquidViscosity / LiquidDensity / DiffusionCoefficient

End Function

Function GetStantonNumber(IonicTransportCoefficient As Double, BedEffectiveContactTime As Double, BedPorosity As Double, ParticleRadius As Double) As Double
   'Input:
   '   IonicTransportCoefficient in cm/s
   '   BedEffectiveContactTime in s
   '   BedPorosity in (-)
   '   ParticleRadius in cm
   'Output:
   '   StantonNumber in (-)

    GetStantonNumber = IonicTransportCoefficient * BedEffectiveContactTime * (1# - BedPorosity) / ParticleRadius / BedPorosity

End Function

Function GetSuperficialVelocity(InletFlowrate As Double, BedArea As Double) As Double
   'Input:
   '   InletFlowrate in m3/s
   '   BedArea in m2
   'Output:
   '   SuperficialVelocity in cm/s

   GetSuperficialVelocity = (InletFlowrate / BedArea) * 100

End Function

Function GetSurfaceDistributionParameter(ApparentDensity As Double, TotalResinCapacity, BedPorosity As Double, SumOfInitialEquivalentConcs) As Double
   'Input:
   '   ApparentDensity in g/cm3
   '   TotalResinCapacity in meq/g
   '   BedPorosity in (-)
   '   SumOfInitialEquivalentConcs in meq/L
   'Output:
   '   SurfaceDistributionParameter in (-)

    GetSurfaceDistributionParameter = (ApparentDensity * 1000#) * TotalResinCapacity * (1# - BedPorosity) / BedPorosity / SumOfInitialEquivalentConcs

End Function

Function GetTotalDistributionParameter(SurfaceDistributionParameter As Double, PoreDistributionParameter As Double) As Double
   'Input:
   '   SurfaceDistributionParameter as (-)
   '   PoreDistributionParameter as (-)
   'Output:
   '   TotalDistributionParameter as (-)

    GetTotalDistributionParameter = SurfaceDistributionParameter + PoreDistributionParameter

End Function

Sub UpdateDimensionlessGroupAllIons()
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

Sub UpdateKineticParametersAllIons()
    Dim i As Integer

    EditingAnion = False
    EditingCation = True
    For i = 1 To NowProj.NumberOfCations
        NumberOfIonToEdit = i
        Call CalculateKineticParameters
    Next i
    EditingCation = False
    EditingAnion = True
    For i = 1 To NowProj.NumberOfAnions
        NumberOfIonToEdit = i
        Call CalculateKineticParameters
    Next i
    EditingAnion = False

End Sub

