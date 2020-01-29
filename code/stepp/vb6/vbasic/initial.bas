Attribute VB_Name = "InitialMod"
'This module contains routines used for initialization of
'various items

Sub InitializeBIPdbHierarchy()

' "UNIFAC Binary Interaction Parameter Database Hierarchy for Activity Coefficient"
       BIP_dbHierarchy.ActivityCoefficient(1) = 3  '  Environmental
       BIP_dbHierarchy.ActivityCoefficient(2) = 1  '  Original UNIFAC VLE
       BIP_dbHierarchy.ActivityCoefficient(3) = 2  '  UNIFAC LLE
       
' "UNIFAC Binary Interaction Parameter Database Hierarchy for Aqueous Solubility"
       BIP_dbHierarchy.AqueousSolubility(1) = 2  '    UNIFAC LLE
       BIP_dbHierarchy.AqueousSolubility(2) = 3  '    Environmental
       BIP_dbHierarchy.AqueousSolubility(3) = 1  '    Original UNIFAC VLE
       
' "UNIFAC Binary Interaction Parameter Database Hierarchy for Octanol Water Partition Coefficient"
       BIP_dbHierarchy.OctWaterPartCoeff(1) = 2  '    UNIFAC LLE
       BIP_dbHierarchy.OctWaterPartCoeff(2) = 1  '    Original UNIFAC VLE

End Sub

Sub InitializeCurrentSelections()

    phprop.VaporPressure.CurrentSelection.choice = 0
    phprop.ActivityCoefficient.CurrentSelection.choice = 0
    phprop.HenrysConstant.CurrentSelection.choice = 0
    phprop.MolecularWeight.CurrentSelection.choice = 0
    phprop.BoilingPoint.CurrentSelection.choice = 0
    phprop.LiquidDensity.CurrentSelection.choice = 0
    phprop.MolarVolume.operatingT.CurrentSelection.choice = 0
    phprop.MolarVolume.BoilingPoint.CurrentSelection.choice = 0
    phprop.RefractiveIndex.CurrentSelection.choice = 0
    phprop.AqueousSolubility.CurrentSelection.choice = 0
    phprop.OctWaterPartCoeff.CurrentSelection.choice = 0
    phprop.LiquidDiffusivity.CurrentSelection.choice = 0
    phprop.GasDiffusivity.CurrentSelection.choice = 0
    phprop.WaterDensity.CurrentSelection.choice = 0
    phprop.WaterViscosity.CurrentSelection.choice = 0
    phprop.WaterSurfaceTension.CurrentSelection.choice = 0
    phprop.AirDensity.CurrentSelection.choice = 0
    phprop.AirViscosity.CurrentSelection.choice = 0

End Sub

Sub InitializeHierarchy()

    hie.VaporPressure(1).hierarchy = 3
    hie.VaporPressure(1).source = "Database"
    hie.VaporPressure(2).hierarchy = 4
    hie.VaporPressure(2).source = "Input"
    
    hie.ActivityCoefficient(1).hierarchy = 5
    
    hie.HenrysConstant(1).hierarchy = 7
    hie.HenrysConstant(1).source = "Regression of Data Pts"
    hie.HenrysConstant(2).hierarchy = 8
    hie.HenrysConstant(2).source = "Fit of UNIFAC w/Data Pt"
    hie.HenrysConstant(3).hierarchy = 9
    hie.HenrysConstant(3).source = "UNIFAC at Operating T"
    hie.HenrysConstant(4).hierarchy = 10
    hie.HenrysConstant(4).source = "Database"
    hie.HenrysConstant(5).hierarchy = 11
    hie.HenrysConstant(5).source = "UNIFAC at Database T's"
    hie.HenrysConstant(6).hierarchy = 12
    hie.HenrysConstant(6).source = "Input"
    
    hie.MolecularWeight(1).hierarchy = 13
    hie.MolecularWeight(1).source = "Database"
    hie.MolecularWeight(2).hierarchy = 14
    hie.MolecularWeight(2).source = "Group Contribution"
    hie.MolecularWeight(3).hierarchy = 15
    hie.MolecularWeight(3).source = "Input"
    
    hie.BoilingPoint(1).hierarchy = 16
    hie.BoilingPoint(1).source = "Database"
    hie.BoilingPoint(2).hierarchy = 17
    hie.BoilingPoint(2).source = "Input"
    
    hie.MolarVolumeBoilingPoint(1).hierarchy = 21
    hie.MolarVolumeBoilingPoint(1).source = "Group Contribution"
    hie.MolarVolumeBoilingPoint(2).hierarchy = 22
    hie.MolarVolumeBoilingPoint(2).source = "Input"
    
    hie.LiquidDensity(1).hierarchy = 18
    hie.LiquidDensity(1).source = "Database"
    hie.LiquidDensity(2).hierarchy = 19
    hie.LiquidDensity(2).source = "Group Contribution"
    hie.LiquidDensity(3).hierarchy = 20
    hie.LiquidDensity(3).source = "Input"
    
    hie.MolarVolumeOperatingT(1).hierarchy = 23
    hie.MolarVolumeOperatingT(1).source = "Database"
    hie.MolarVolumeOperatingT(2).hierarchy = 24
    hie.MolarVolumeOperatingT(2).source = "Group Contribution"
    hie.MolarVolumeOperatingT(3).hierarchy = 25
    hie.MolarVolumeOperatingT(3).source = "Input"
    
    hie.RefractiveIndex(1).hierarchy = 26
    hie.RefractiveIndex(1).source = "Database"
    hie.RefractiveIndex(2).hierarchy = 27
    hie.RefractiveIndex(2).source = "Input"
    
    hie.AqueousSolubility(1).hierarchy = 28
    hie.AqueousSolubility(1).source = "Fit"
    hie.AqueousSolubility(2).hierarchy = 29
    hie.AqueousSolubility(2).source = "UNIFAC at Operating T"
    hie.AqueousSolubility(3).hierarchy = 30
    hie.AqueousSolubility(3).source = "Database"
    hie.AqueousSolubility(4).hierarchy = 31
    hie.AqueousSolubility(4).source = "UNIFAC at Database T"
    hie.AqueousSolubility(5).hierarchy = 32
    hie.AqueousSolubility(5).source = "Input"
    
    hie.OctWaterPartCoeff(1).hierarchy = 35
    hie.OctWaterPartCoeff(1).source = "UNIFAC at Operating T"
    hie.OctWaterPartCoeff(2).hierarchy = 33
    hie.OctWaterPartCoeff(2).source = "Database"
    hie.OctWaterPartCoeff(3).hierarchy = 34
    hie.OctWaterPartCoeff(3).source = "UNIFAC at Database T"
    hie.OctWaterPartCoeff(4).hierarchy = 36
    hie.OctWaterPartCoeff(4).source = "Input"
    
    hie.LiquidDiffusivityMWTlt1000(1).hierarchy = 38
    hie.LiquidDiffusivityMWTlt1000(1).source = "Hayduk & Laudie"
    hie.LiquidDiffusivityMWTlt1000(2).hierarchy = 39
    hie.LiquidDiffusivityMWTlt1000(2).source = "Wilke-Chang"
    hie.LiquidDiffusivityMWTlt1000(3).hierarchy = 37
    hie.LiquidDiffusivityMWTlt1000(3).source = "Polson"
    hie.LiquidDiffusivityMWTlt1000(4).hierarchy = 40
    hie.LiquidDiffusivityMWTlt1000(4).source = "Input"
    
    hie.LiquidDiffusivityMWTgt1000(1).hierarchy = 37
    hie.LiquidDiffusivityMWTgt1000(1).source = "Polson"
    hie.LiquidDiffusivityMWTgt1000(2).hierarchy = 38
    hie.LiquidDiffusivityMWTgt1000(2).source = "Hayduk & Laudie"
    hie.LiquidDiffusivityMWTgt1000(3).hierarchy = 39
    hie.LiquidDiffusivityMWTgt1000(3).source = "Wilke-Chang"
    hie.LiquidDiffusivityMWTgt1000(4).hierarchy = 40
    hie.LiquidDiffusivityMWTgt1000(4).source = "Input"
    
    hie.GasDiffusivity(1).hierarchy = 41
    hie.GasDiffusivity(1).source = "Wilke-Lee"
    hie.GasDiffusivity(2).hierarchy = 42
    hie.GasDiffusivity(2).source = "Input"
    
    hie.WaterDensity(1).hierarchy = 43
    hie.WaterDensity(1).source = "Correlation"
    hie.WaterDensity(2).hierarchy = 44
    hie.WaterDensity(2).source = "Input"
    
    hie.WaterViscosity(1).hierarchy = 45
    hie.WaterViscosity(1).source = "Correlation"
    hie.WaterViscosity(2).hierarchy = 46
    hie.WaterViscosity(2).source = "Input"
    
    hie.WaterSurfaceTension(1).hierarchy = 47
    hie.WaterSurfaceTension(1).source = "Correlation"
    hie.WaterSurfaceTension(2).hierarchy = 48
    hie.WaterSurfaceTension(2).source = "Input"
    
    hie.AirDensity(1).hierarchy = 49
    hie.AirDensity(1).source = "Correlation"
    hie.AirDensity(2).hierarchy = 50
    hie.AirDensity(2).source = "Input"
    
    hie.AirViscosity(1).hierarchy = 51
    hie.AirViscosity(1).source = "Correlation"
    hie.AirViscosity(2).hierarchy = 52
    hie.AirViscosity(2).source = "Input"

End Sub

Sub InitializeHilights()
    hilight.VaporPressure.PreviousIndex = -1
    hilight.ActivityCoefficient.PreviousIndex = -1
    hilight.HenrysConstant.PreviousIndex = -1
    hilight.MolecularWeight.PreviousIndex = -1
    hilight.BoilingPoint.PreviousIndex = -1
    hilight.LiquidDensity.PreviousIndex = -1
    hilight.MolarVolumeOperatingT.PreviousIndex = -1
    hilight.MolarVolumeBoilingPoint.PreviousIndex = -1
    hilight.RefractiveIndex.PreviousIndex = -1
    hilight.AqueousSolubility.PreviousIndex = -1
    hilight.OctWaterPartCoeff.PreviousIndex = -1
    hilight.LiquidDiffusivity.PreviousIndex = -1
    hilight.GasDiffusivity.PreviousIndex = -1
    hilight.WaterDensity.PreviousIndex = -1
    hilight.WaterViscosity.PreviousIndex = -1
    hilight.WaterSurfaceTension.PreviousIndex = -1
    hilight.AirDensity.PreviousIndex = -1
    hilight.AirViscosity.PreviousIndex = -1

End Sub

Sub InitializePROPandHAVEAVAILABLEArrays()

    For I = 1 To NUMBER_OF_PROPERTIES
        HaveProperty(I) = False
    Next I

    For I = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
        PROPAVAILABLE(I) = False
    Next I

End Sub

Sub InitializeUserInputs()
'*** This subroutine will initialize the input variables for all
'*** the properties.  It will be needed to successfully transfer from
'*** chemical to chemical

    '*** Vapor Pressure
    phprop.VaporPressure.input.Value = -1#
    phprop.VaporPressure.input.temperature = -1E+25

    '*** Activity Coefficient
    phprop.ActivityCoefficient.input.Value = -1#
    phprop.ActivityCoefficient.input.temperature = -1E+25

    '*** Henry's Constant
    phprop.HenrysConstant.input.Value = -1#
    phprop.HenrysConstant.input.temperature = -1E+25

    '*** Molecular Weight
    phprop.MolecularWeight.input.Value = -1#

    '*** Normal Boiling Point
    phprop.BoilingPoint.input.Value = -1E+25

    '*** Liquid Density
    phprop.LiquidDensity.input.Value = -1#
    phprop.LiquidDensity.input.temperature = -1E+25

    '*** Molar Volume at Operating Temperature
    phprop.MolarVolume.operatingT.input.Value = -1#
    phprop.MolarVolume.operatingT.input.temperature = -1E+25

    '*** Molar Volume at Normal Boiling Point
    phprop.MolarVolume.BoilingPoint.input.Value = -1#
    phprop.MolarVolume.BoilingPoint.input.temperature = -1E+25

    '*** Refractive Index
    phprop.RefractiveIndex.input.Value = -1#

    '*** Aqueous Solubility
    phprop.AqueousSolubility.input.Value = -1#
    phprop.AqueousSolubility.input.temperature = -1E+25

    '*** Octanol Water Partition Coefficient
    phprop.OctWaterPartCoeff.input.Value = -1E+25
    phprop.OctWaterPartCoeff.input.temperature = -1E+25

    '*** Liquid Diffusivity
    phprop.LiquidDiffusivity.input.Value = -1#
    phprop.LiquidDiffusivity.input.temperature = -1E+25

    '*** Gas Diffusivity
    phprop.GasDiffusivity.input.Value = -1#
    phprop.GasDiffusivity.input.temperature = -1E+25

    '*** Water Density
    phprop.WaterDensity.input.Value = -1#
    phprop.WaterDensity.input.temperature = -1E+25

    '*** Water Viscosity
    phprop.WaterViscosity.input.Value = -1#
    phprop.WaterViscosity.input.temperature = -1E+25

    '*** Water Surface Tension
    phprop.WaterSurfaceTension.input.Value = -1#
    phprop.WaterSurfaceTension.input.temperature = -1E+25

    '*** Air Density
    phprop.AirDensity.input.Value = -1#
    phprop.AirDensity.input.temperature = -1E+25

    '*** Air Viscosity
    phprop.AirViscosity.input.Value = -1#
    phprop.AirViscosity.input.temperature = -1E+25

End Sub

