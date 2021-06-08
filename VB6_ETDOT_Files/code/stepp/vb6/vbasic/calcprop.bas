Attribute VB_Name = "CalcPropMod"
'*** This module is used to calculate the properties of
'*** StEPP.  The Calculation of each individual property
'*** will be done in a separate subroutine.

'   *** Array needed by MWTCALL and AQSCALL
'    Global XMW(1 To ND) As Double

Sub CalculateActivityCoefficient()
    Dim msg As String
    

'      *******************************************************
'      *                                                     *
'      *      Infinite Dilution Activity Coefficient         *
'      *                                                     *
'      *******************************************************

       If phprop.MaximumUnifacGroups > 0 Then

CalculateUNIFACActCoeffOperatingT:

          phprop.ActivityCoefficient.UNIFAC.Value = 0#
          phprop.ActivityCoefficient.UNIFAC.error = 0

          On Error GoTo ActivityCoefficientUNIFACError
          Call ACCALL( _
              phprop.ActivityCoefficient.UNIFAC.Value, _
              phprop.ActivityCoefficient.UNIFAC.Source.short, _
              phprop.ActivityCoefficient.UNIFAC.Source.long, _
              phprop.ActivityCoefficient.UNIFAC.error, _
              phprop.ActivityCoefficient.UNIFAC.temperature, _
              phprop.OperatingTemperature, _
              FGRPErrorFlag, _
              phprop.MaximumUnifacGroups, _
              phprop.MS(1, 1, 1), _
              phprop.ActivityCoefficient.BinaryInteractionParameterDatabase)

          If phprop.ActivityCoefficient.UNIFAC.error < 0 Then 'Error calculating activity coefficient with this particular UNIFAC parameter set
             phprop.ActivityCoefficient.BinaryInteractionParameterDBAvailable(phprop.ActivityCoefficient.BinaryInteractionParameterDatabase) = False
             If UserSelectedTheUnifacBIPDBActCoeff Then
                phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB
                MsgBox "Selected database not available to calculate activity coefficient for this compound.  Returning to Original Choice", MB_ICONSTOP, "Data Not Available"
                Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB - 1
                GoTo CalculateUNIFACActCoeffOperatingT
             End If

             Select Case phprop.ActivityCoefficient.BinaryInteractionParameterDatabase
                Case BIP_dbHierarchy.ActivityCoefficient(1)
                   phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = BIP_dbHierarchy.ActivityCoefficient(2)
                   GoTo CalculateUNIFACActCoeffOperatingT
                Case BIP_dbHierarchy.ActivityCoefficient(2)
                   phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = BIP_dbHierarchy.ActivityCoefficient(3)
                   GoTo CalculateUNIFACActCoeffOperatingT
                Case BIP_dbHierarchy.ActivityCoefficient(3)
                    phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = 0
             End Select
          End If
          If phprop.ActivityCoefficient.UNIFAC.error < 0 Then
             phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = 0
          End If

          
       Else
          phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = 0
          phprop.ActivityCoefficient.UNIFAC.error = -36
       End If

       If phprop.ActivityCoefficient.UNIFAC.error >= 0 Then
          PROPAVAILABLE(ACTIVITY_COEFFICIENT_UNIFAC) = True
       Else
          If phprop.ActivityCoefficient.CurrentSelection.choice = ACTIVITY_COEFFICIENT_UNIFAC Then
             phprop.ActivityCoefficient.CurrentSelection.choice = 0
             Infinite_dilution_form!lblSourceLabel(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(ACTIVITY_COEFFICIENT_UNIFAC) = False
       End If

      Call DisplayActivityCoefficient

      Exit Sub

ActivityCoefficientUNIFACError:
      msg = "Error in the FORTRAN routines while calculating Activity Coefficient from UNIFAC!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.ActivityCoefficient.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateAirDensity()

'        ********************************************************
'        *                                                      *
'        *              Air Density                             *
'        *                                                      *
'        ********************************************************

        
        On Error GoTo AirDensityCorrelationError
        Call AIRDENS(phprop.AirDensity.correlation.Value, phprop.OperatingTemperature, phprop.OperatingPressure, phprop.AirDensity.correlation.error, phprop.AirDensity.correlation.Source.short, phprop.AirDensity.correlation.Source.long, phprop.AirDensity.correlation.temperature)

      If phprop.AirDensity.correlation.error >= 0 Then
         PROPAVAILABLE(AIR_DENSITY_CORRELATION) = True
      Else
         If phprop.AirDensity.CurrentSelection.choice = AIR_DENSITY_CORRELATION Then
            phprop.AirDensity.CurrentSelection.choice = 0
            frmAirDensity!lblSource(0).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(AIR_DENSITY_CORRELATION) = False
      End If

       Call DisplayAirDensity

      Exit Sub

AirDensityCorrelationError:
      msg = "Error in the FORTRAN routines while calculating Air Density from Correlation!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.AirDensity.correlation.error = -200
      Resume Next

End Sub

Sub CalculateAirViscosity()

'       *******************************************************
'       *                                                     *
'       *              Air Viscosity                          *
'       *                                                     *
'       *******************************************************

      On Error GoTo AirViscosityCorrelationError
      Call AIRVISC(phprop.AirViscosity.correlation.Value, phprop.OperatingTemperature, phprop.AirViscosity.correlation.error, phprop.AirViscosity.correlation.Source.short, phprop.AirViscosity.correlation.Source.long, phprop.AirViscosity.correlation.temperature)

      If phprop.AirViscosity.correlation.error >= 0 Then
         PROPAVAILABLE(AIR_VISCOSITY_CORRELATION) = True
      Else
         If phprop.AirViscosity.CurrentSelection.choice = AIR_VISCOSITY_CORRELATION Then
            phprop.AirViscosity.CurrentSelection.choice = 0
            frmAirViscosity!lblSource(0).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(AIR_VISCOSITY_CORRELATION) = False
      End If

       Call DisplayAirViscosity
      
      Exit Sub

AirViscosityCorrelationError:
      msg = "Error in the FORTRAN routines while calculating Air Viscosity from Correlation!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.AirViscosity.correlation.error = -200
      Resume Next

End Sub

Sub CalculateAqueousSolubility()
    Dim i As Integer
    Dim J As Integer
    Dim K As Integer


'        ******************************************************
'        *                                                    *
'        *               Aqueous Solubility                   *
'        *                                                    *
'        ******************************************************


'   /***** VALUE FROM DATABASE */

        If phprop.AqueousSolubility.database.Value < 0 Then
           phprop.AqueousSolubility.database.error = -22
        Else
           phprop.AqueousSolubility.database.error = 0
        End If

       If phprop.AqueousSolubility.database.error >= 0 Then
          PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) = True
       Else
          If phprop.AqueousSolubility.CurrentSelection.choice = AQUEOUS_SOLUBILITY_DATABASE Then
             phprop.AqueousSolubility.CurrentSelection.choice = 0
             aqsol_form!lblSource(2).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) = False
       End If


'   /***** Value from UNIFAC at operating temperature */

       If phprop.MaximumUnifacGroups > 0 Then
CalculateUNIFACAqSolOperatingT:

          phprop.AqueousSolubility.operatingT.UNIFAC.Value = 0#
          phprop.AqueousSolubility.operatingT.UNIFAC.error = 0

          On Error GoTo AqueousSolubilityUNIFACopTError
          Call AQSCALL(phprop.AqueousSolubility.operatingT.UNIFAC.Value, phprop.AqueousSolubility.operatingT.UNIFAC.Source.short, phprop.AqueousSolubility.operatingT.UNIFAC.Source.long, phprop.AqueousSolubility.operatingT.UNIFAC.error, phprop.AqueousSolubility.operatingT.UNIFAC.temperature, phprop.OperatingTemperature, phprop.MaximumUnifacGroups, phprop.MS(1, 1, 1), phprop.XMW(1), phprop.AqueousSolubility.BinaryInteractionParameterDatabase)

          If phprop.AqueousSolubility.operatingT.UNIFAC.error < 0 Then 'Error calculating solubility with this particular UNIFAC parameter set
             phprop.AqueousSolubility.BinaryInteractionParameterDBAvailable(phprop.AqueousSolubility.BinaryInteractionParameterDatabase) = False
             If UserSelectedTheUnifacBIPDBAqSol Then
                phprop.AqueousSolubility.BinaryInteractionParameterDatabase = phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB
                MsgBox "Selected UNIFAC database not available to calculate aqueous solubility for this compound.  Returning to Original Choice", MB_ICONSTOP, "Data Not Available"
                aqsol_form!cboUNIFACParameterSet.ListIndex = phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB - 1
                GoTo CalculateUNIFACAqSolOperatingT
             End If

             Select Case phprop.AqueousSolubility.BinaryInteractionParameterDatabase
                Case BIP_dbHierarchy.AqueousSolubility(1)
                   phprop.AqueousSolubility.BinaryInteractionParameterDatabase = BIP_dbHierarchy.AqueousSolubility(2)
                   GoTo CalculateUNIFACAqSolOperatingT
                Case BIP_dbHierarchy.AqueousSolubility(2)
                   phprop.AqueousSolubility.BinaryInteractionParameterDatabase = BIP_dbHierarchy.AqueousSolubility(3)
                   GoTo CalculateUNIFACAqSolOperatingT
                Case BIP_dbHierarchy.AqueousSolubility(3)
                    phprop.AqueousSolubility.BinaryInteractionParameterDatabase = 0
             End Select
          End If
          If phprop.AqueousSolubility.operatingT.UNIFAC.error < 0 Then
             phprop.AqueousSolubility.BinaryInteractionParameterDatabase = 0
          End If
       Else
          phprop.AqueousSolubility.BinaryInteractionParameterDatabase = 0
          phprop.AqueousSolubility.operatingT.UNIFAC.error = -36
       End If

       If phprop.AqueousSolubility.operatingT.UNIFAC.error >= 0 Then
          PROPAVAILABLE(AQUEOUS_SOLUBILITY_OPT_UNIFAC) = True
       Else
          If phprop.AqueousSolubility.CurrentSelection.choice = AQUEOUS_SOLUBILITY_OPT_UNIFAC Then
             phprop.AqueousSolubility.CurrentSelection.choice = 0
             aqsol_form!lblSource(1).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(AQUEOUS_SOLUBILITY_OPT_UNIFAC) = False
       End If


'      ********* Value from Unifac at Database Temperature *
      
      If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) Then

         If phprop.MaximumUnifacGroups > 0 Then
            On Error GoTo AqueousSolubilityUNIFACdbTError
            Call AQSCALL(phprop.AqueousSolubility.UNIFAC.Value, phprop.AqueousSolubility.UNIFAC.Source.short, phprop.AqueousSolubility.UNIFAC.Source.long, phprop.AqueousSolubility.UNIFAC.error, phprop.AqueousSolubility.UNIFAC.temperature, phprop.AqueousSolubility.database.temperature, phprop.MaximumUnifacGroups, phprop.MS(1, 1, 1), phprop.XMW(1), phprop.AqueousSolubility.BinaryInteractionParameterDatabase)
       Else
          phprop.AqueousSolubility.UNIFAC.error = -36
       End If
                                                
         If phprop.AqueousSolubility.UNIFAC.error >= 0 Then
            PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC) = True
         Else
            If phprop.AqueousSolubility.CurrentSelection.choice = AQUEOUS_SOLUBILITY_DBT_UNIFAC Then
               phprop.AqueousSolubility.CurrentSelection.choice = 0
               aqsol_form!lblSource(3).BackColor = &HC0C0C0
            End If
            PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC) = False
         End If

      Else
         phprop.AqueousSolubility.UNIFAC.error = -44
         If phprop.AqueousSolubility.CurrentSelection.choice = AQUEOUS_SOLUBILITY_DBT_UNIFAC Then
            phprop.AqueousSolubility.CurrentSelection.choice = 0
            aqsol_form!lblSource(3).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC) = False
      End If


'      **** Value from fit of UNIFAC with a data point

      If (PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) And PROPAVAILABLE(AQUEOUS_SOLUBILITY_OPT_UNIFAC) And PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC)) Then
          phprop.AqueousSolubility.fit.UNIFAC.error = 0
          On Error GoTo AqueousSolubilityUNIFACfitError
             Call AQSFIT(phprop.AqueousSolubility.fit.UNIFAC.Value, phprop.AqueousSolubility.fit.UNIFAC.Source.short, phprop.AqueousSolubility.fit.UNIFAC.Source.long, phprop.AqueousSolubility.fit.UNIFAC.error, phprop.AqueousSolubility.fit.UNIFAC.temperature, phprop.AqueousSolubility.UNIFAC.Value, phprop.AqueousSolubility.UNIFAC.temperature, phprop.AqueousSolubility.operatingT.UNIFAC.Value, phprop.AqueousSolubility.database.Value, phprop.AqueousSolubility.database.temperature, phprop.OperatingTemperature)

          'Fit routine may produce a negative solubility so if this happens, make solubility unavailable from the fit (error = -45)
          If (phprop.AqueousSolubility.fit.UNIFAC.Value < 0#) Then
             phprop.AqueousSolubility.fit.UNIFAC.error = -45
             If phprop.AqueousSolubility.CurrentSelection.choice = AQUEOUS_SOLUBILITY_FIT Then
                phprop.AqueousSolubility.CurrentSelection.choice = 0
                aqsol_form!lblSource(0).BackColor = &HC0C0C0
                aqsol_form!lblSource(0).ForeColor = &H80000008
                hilight.AqueousSolubility.PreviousIndex = -1
             End If
          End If
      Else
         phprop.AqueousSolubility.fit.UNIFAC.error = -45
      End If

       If phprop.AqueousSolubility.fit.UNIFAC.error >= 0 Then
          PROPAVAILABLE(AQUEOUS_SOLUBILITY_FIT) = True
       Else
          If phprop.AqueousSolubility.CurrentSelection.choice = AQUEOUS_SOLUBILITY_FIT Then
             phprop.AqueousSolubility.CurrentSelection.choice = 0
             aqsol_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(AQUEOUS_SOLUBILITY_FIT) = False
       End If

      Call DisplayAqueousSolubility

      Exit Sub

AqueousSolubilityUNIFACopTError:
      msg = "Error in the FORTRAN routines while calculating Aqueous Solubility from UNIFAC at Operating T!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.AqueousSolubility.operatingT.UNIFAC.error = -200
      Resume Next

AqueousSolubilityUNIFACdbTError:
      msg = "Error in the FORTRAN routines while calculating Aqueous Solubility from UNIFAC at Database T!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.AqueousSolubility.UNIFAC.error = -200
      Resume Next

AqueousSolubilityUNIFACfitError:
      msg = "Error in the FORTRAN routines while calculating Aqueous Solubility from UNIFAC Fit with a Data Point!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.AqueousSolubility.fit.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateBoilingPoint()

'       ********************************************************
'       *                                                      *
'       *             Normal Boiling Point                     *
'       *                                                      *
'       ********************************************************

'      /*  Note:  This value is only available in the database, but
'                 I put this note here so we would know it is a
'                 property that fits here in our structure.  No
'                 UNIFAC calculations are needed for it. */

      If (phprop.BoilingPoint.database.Source.short = -1) Then
          phprop.BoilingPoint.database.error = -16
          If phprop.BoilingPoint.CurrentSelection.choice = BOILING_POINT_DATABASE Then
             phprop.BoilingPoint.CurrentSelection.choice = 0
             nbp_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(BOILING_POINT_DATABASE) = False
      Else
          PROPAVAILABLE(BOILING_POINT_DATABASE) = True
          phprop.BoilingPoint.database.error = 0
      End If

      Call DisplayBoilingPoint

End Sub

Sub CalculateGasDiffusivity()

'     *********************************************************
'     *                                                       *
'     *              Gas Diffusivity                          *
'     *                                                       *
'     *********************************************************

'      ******  Value from the Wilke-Lee Modification of the
'      ******  Hirschfelder-Bird-Spotz Method

       If HaveProperty(MOLAR_VOLUME_BOILING_POINT) And HaveProperty(BOILING_POINT) And HaveProperty(MOLECULAR_WEIGHT) Then

          On Error GoTo GasDiffusivityWilkeLeeError
          Call DIFGWL(phprop.GasDiffusivity.wilkeLee.Value, phprop.MolecularWeight.CurrentSelection.Value, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value, phprop.BoilingPoint.CurrentSelection.Value, phprop.OperatingTemperature, phprop.OperatingPressure, phprop.GasDiffusivity.wilkeLee.error, phprop.GasDiffusivity.wilkeLee.Source.short, phprop.GasDiffusivity.wilkeLee.Source.long, phprop.GasDiffusivity.wilkeLee.temperature)

          If phprop.GasDiffusivity.wilkeLee.error >= 0 Then
             PROPAVAILABLE(GAS_DIFFUSIVITY_WILKELEE) = True
          Else
             If phprop.GasDiffusivity.CurrentSelection.choice = GAS_DIFFUSIVITY_WILKELEE Then
                phprop.GasDiffusivity.CurrentSelection.choice = 0
                gas_diff_form!lblSource(0).BackColor = &HC0C0C0
             End If
             PROPAVAILABLE(GAS_DIFFUSIVITY_WILKELEE) = False
          End If

       Else
          If HaveProperty(BOILING_POINT) And HaveProperty(MOLECULAR_WEIGHT) Then
             phprop.GasDiffusivity.wilkeLee.error = -34
          ElseIf HaveProperty(MOLAR_VOLUME_BOILING_POINT) And HaveProperty(MOLECULAR_WEIGHT) Then
             phprop.GasDiffusivity.wilkeLee.error = -48
          ElseIf HaveProperty(MOLAR_VOLUME_BOILING_POINT) And HaveProperty(BOILING_POINT) Then
             phprop.GasDiffusivity.wilkeLee.error = -49
          ElseIf HaveProperty(MOLECULAR_WEIGHT) Then
             phprop.GasDiffusivity.wilkeLee.error = -50
          ElseIf HaveProperty(BOILING_POINT) Then
             phprop.GasDiffusivity.wilkeLee.error = -51
          ElseIf HaveProperty(MOLAR_VOLUME_BOILING_POINT) Then
             phprop.GasDiffusivity.wilkeLee.error = -52
          Else
             phprop.GasDiffusivity.wilkeLee.error = -53
          End If
          If phprop.BoilingPoint.CurrentSelection.choice = BOILING_POINT_DATABASE Then
             phprop.AqueousSolubility.CurrentSelection.choice = 0
             nbp_form!lblSource(0).BackColor = &HC0C0C0
          End If
          If phprop.GasDiffusivity.CurrentSelection.choice = GAS_DIFFUSIVITY_WILKELEE Then
             phprop.GasDiffusivity.CurrentSelection.choice = 0
             gas_diff_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(GAS_DIFFUSIVITY_WILKELEE) = False
       End If

       Call DisplayGasDiffusivity

      Exit Sub

GasDiffusivityWilkeLeeError:
      msg = "Error in the FORTRAN routines while calculating Gas Diffusivity from Wilke-Lee Modification of Hirschfelder-Bird-Spotz Method!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.GasDiffusivity.wilkeLee.error = -200
      Resume Next

End Sub

Sub CalculateHenrysConstant()
    Static HenrysConstantDatabaseVal(1 To Maxchemical)  As Double
    Static HenrysConstantDatabaseTemp(1 To Maxchemical)  As Double

    Static HenrysConstantUnifacVal(1 To Maxchemical)  As Double
    Static HenrysConstantUnifacShortSrc(1 To Maxchemical) As Long
    Static HenrysConstantUnifacLongSrc(1 To Maxchemical) As Long
    Static HenrysConstantUnifacErr(1 To Maxchemical) As Long
    Static HenrysConstantUnifacTemp(1 To Maxchemical) As Double

'   *** Declare values to be used to calculate Henry's constants
'   *** corresponding to database temperatures.  Some of these values
'   *** are not used except to prevent overwriting data.

    Dim VaporPressureValHC As Double
    Dim VaporPressureShortSrcHC As Long
    Dim VaporPressureLongSrcHC As Long
    Dim VaporPressureErrHC As Long
    Dim VaporPressureTempHC As Double

    Dim ActivityCoeffValHC As Double
    Dim ActivityCoeffShortSrcHC As Long
    Dim ActivityCoeffLongSrcHC As Long
    Dim ActivityCoeffErrHC As Long
    Dim ActivityCoeffTempHC As Double

    Dim hc_unifac_value As String * 40
    Dim hc_unifac_temp As String
    Dim hc_string As String

'   *** Set arrays for passing to FORTRAN routines

    For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants

        HenrysConstantDatabaseVal(i) = phprop.HenrysConstant.database(i).Value
        HenrysConstantDatabaseTemp(i) = phprop.HenrysConstant.database(i).temperature

    Next i

'       *********************************************************
'       *                                                       *
'       *                Henry's Constant                       *
'       *                                                       *
'       *********************************************************


'   *** Henry's Constants from Database:
'   *** Find T in Database Closest to Operating Temperature

    If phprop.HenrysConstant.NumberOfDatabaseHenrysConstants = 0 Then
       phprop.HenrysConstant.chosenDatabaseIndex = 0
    ElseIf phprop.HenrysConstant.NumberOfDatabaseHenrysConstants = 1 Then
       phprop.HenrysConstant.chosenDatabaseIndex = 1
    Else
       Call GetClosestHCDatabaseT
    End If


'   *** Set Activity coefficient and vapor pressure to values
'   *** calculated above.  Eventually, what these are set to will
'   *** have to take the hierarchy into account

       If (HaveProperty(ACTIVITY_COEFFICIENT) And HaveProperty(VAPOR_PRESSURE)) Then
          
'          /* HC1CALL:  Find UNIFAC Henry's constant at operating T */
          
          On Error GoTo HenrysConstantUNIFACopTError
          Call HC1CALL( _
              phprop.HenrysConstant.operatingT.UNIFAC.Value, _
              phprop.HenrysConstant.operatingT.UNIFAC.Source.short, _
              phprop.HenrysConstant.operatingT.UNIFAC.Source.long, _
              phprop.HenrysConstant.operatingT.UNIFAC.error, _
              phprop.HenrysConstant.operatingT.UNIFAC.temperature, _
              phprop.OperatingTemperature, _
              phprop.ActivityCoefficient.CurrentSelection.Value, _
              phprop.VaporPressure.CurrentSelection.Value)

          If phprop.HenrysConstant.operatingT.UNIFAC.error >= 0 Then
             PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC) = True
          Else
             If phprop.HenrysConstant.CurrentSelection.choice = HENRYS_CONSTANT_OPT_UNIFAC Then
                phprop.HenrysConstant.CurrentSelection.choice = 0
                hc_form!lblSource(2).BackColor = &HC0C0C0
             End If
             PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC) = False
          End If
       Else
          If HaveProperty(ACTIVITY_COEFFICIENT) Then
             phprop.HenrysConstant.operatingT.UNIFAC.error = -38
          ElseIf HaveProperty(VAPOR_PRESSURE) Then
             phprop.HenrysConstant.operatingT.UNIFAC.error = -39
          Else
             phprop.HenrysConstant.operatingT.UNIFAC.error = -40
          End If
          If phprop.HenrysConstant.CurrentSelection.choice = HENRYS_CONSTANT_OPT_UNIFAC Then
             phprop.HenrysConstant.CurrentSelection.choice = 0
             hc_form!lblSource(2).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC) = False
       End If


'  HC2CALL:        Find Henry's Constant at operating T from linear
'                  regression  on data points in database.  This can
'                  only be done if more than one data point is
'                  available in the database.

       On Error GoTo HenrysConstantRegressionError
       Call HC2CALL(phprop.HenrysConstant.regress.Value, phprop.HenrysConstant.regress.Source.short, phprop.HenrysConstant.regress.Source.long, phprop.HenrysConstant.regress.error, phprop.HenrysConstant.regress.temperature, HenrysConstantDatabaseVal(1), HenrysConstantDatabaseTemp(1), phprop.OperatingTemperature, phprop.HenrysConstant.NumberOfDatabaseHenrysConstants)
 
       If phprop.HenrysConstant.regress.error >= 0 Then
          PROPAVAILABLE(HENRYS_CONSTANT_REGRESS) = True
       Else
          If phprop.HenrysConstant.CurrentSelection.choice = HENRYS_CONSTANT_REGRESS Then
             phprop.HenrysConstant.CurrentSelection.choice = 0
             hc_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(HENRYS_CONSTANT_REGRESS) = False
       End If


'       /* Find UNIFAC Values at all temperatures corresponding to values in the database */

      If phprop.MaximumUnifacGroups > 0 Then
         For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
             On Error GoTo VaporPressureErrorHenrysCon
                Call VPRCALL(VaporPressureValHC, phprop.VaporPressure.database.Source.short, VaporPressureLongSrcHC, VaporPressureErrHC, phprop.VaporPressure.database.equation, VaporPressureTempHC, phprop.VaporPressure.database.minimumT, phprop.VaporPressure.database.maximumT, phprop.VaporPressure.database.antoineA, phprop.VaporPressure.database.antoineB, phprop.VaporPressure.database.antoineC, phprop.VaporPressure.database.antoineD, phprop.VaporPressure.database.antoineE, phprop.VaporPressure.database.superfund.Value, phprop.VaporPressure.database.superfund.temperature, HenrysConstantDatabaseTemp(i))
             On Error GoTo ActivityCoeffErrorHenrysCon
                Call ACCALL(ActivityCoeffValHC, ActivityCoeffShortSrcHC, ActivityCoeffLongSrcHC, ActivityCoeffErrHC, ActivityCoeffTempHC, HenrysConstantDatabaseTemp(i), FGRPErrorFlag, phprop.MaximumUnifacGroups, phprop.MS(1, 1, 1), phprop.ActivityCoefficient.BinaryInteractionParameterDatabase)
             If ((VaporPressureErrHC >= 0) And (ActivityCoeffErrHC >= 0)) Then
                On Error GoTo HenrysConstantUNIFACdbTError
                   Call HC1CALL(HenrysConstantUnifacVal(i), HenrysConstantUnifacShortSrc(i), HenrysConstantUnifacLongSrc(i), HenrysConstantUnifacErr(i), HenrysConstantUnifacTemp(i), HenrysConstantDatabaseTemp(i), ActivityCoeffValHC, VaporPressureValHC)

'               *** Set the temporary values to their permanant arrays
                phprop.HenrysConstant.UNIFAC(i).Value = HenrysConstantUnifacVal(i)
                phprop.HenrysConstant.UNIFAC(i).Source.short = HenrysConstantUnifacShortSrc(i)
                phprop.HenrysConstant.UNIFAC(i).Source.long = HenrysConstantUnifacLongSrc(i)
                phprop.HenrysConstant.UNIFAC(i).error = HenrysConstantUnifacErr(i)
                phprop.HenrysConstant.UNIFAC(i).temperature = HenrysConstantUnifacTemp(i)
      
             Else  '*** Correct this error number later
                phprop.HenrysConstant.UNIFAC(i).error = -54
             End If

         Next i
      Else   '*** No Unifac Henry's constants are available
         For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
             phprop.HenrysConstant.UNIFAC(i).error = -36
         Next i
      End If

      If phprop.HenrysConstant.NumberOfDatabaseHenrysConstants = 0 Then
         If phprop.HenrysConstant.CurrentSelection.choice = HENRYS_CONSTANT_DATABASE Then
            phprop.HenrysConstant.CurrentSelection.choice = 0
            hc_form!lblSource(3).BackColor = &HC0C0C0
         End If
         If phprop.HenrysConstant.CurrentSelection.choice = HENRYS_CONSTANT_UNIFAC Then
            phprop.HenrysConstant.CurrentSelection.choice = 0
            hc_form!lblSource(4).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(HENRYS_CONSTANT_DATABASE) = False
         PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) = False
      Else
         If phprop.HenrysConstant.CurrentSelection.choice = HENRYS_CONSTANT_UNIFAC Then
            phprop.HenrysConstant.CurrentSelection.choice = 0
            hc_form!lblSource(4).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(HENRYS_CONSTANT_DATABASE) = True
         PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) = False
         For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
             If phprop.HenrysConstant.UNIFAC(i).error >= 0 Then
                PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) = True
             End If
         Next i
      End If

'   *** Determine index of Unifac Henry's constant closest to operating temperature

      If PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) Then
         phprop.HenrysConstant.chosenUNIFACIndex = phprop.HenrysConstant.chosenDatabaseIndex
         If phprop.HenrysConstant.chosenUNIFACIndex <> 0 Then
            If phprop.HenrysConstant.UNIFAC(phprop.HenrysConstant.chosenUNIFACIndex).error < 0 Then
               If phprop.HenrysConstant.NumberOfDatabaseHenrysConstants = 1 Then
                  phprop.HenrysConstant.chosenUNIFACIndex = 0
               Else
                  Call GetClosestHCUnifacT
               End If
            End If
         End If
      End If


'***     HENFIT:  Find Henry's Constant at operating T from fit of a
'***              single data point in database with UNIFAC values.
'***              This will only be done if at least one data point
'***              is available in the database

        If (PROPAVAILABLE(HENRYS_CONSTANT_DATABASE) And PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) And PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC)) Then
            phprop.HenrysConstant.fit.UNIFAC.error = 0
            On Error GoTo HenrysConstantUNIFACfitError
               Call HENFIT(phprop.HenrysConstant.fit.UNIFAC.Value, phprop.HenrysConstant.fit.UNIFAC.Source.short, phprop.HenrysConstant.fit.UNIFAC.Source.long, phprop.HenrysConstant.fit.UNIFAC.error, phprop.HenrysConstant.fit.UNIFAC.temperature, HenrysConstantDatabaseVal(1), HenrysConstantDatabaseTemp(1), phprop.HenrysConstant.operatingT.UNIFAC.Value, HenrysConstantUnifacVal(1), HenrysConstantUnifacErr(1), phprop.OperatingTemperature, phprop.HenrysConstant.NumberOfDatabaseHenrysConstants)
        Else
             phprop.HenrysConstant.fit.UNIFAC.error = -41
        End If
 
 
       If phprop.HenrysConstant.fit.UNIFAC.error >= 0 Then
          PROPAVAILABLE(HENRYS_CONSTANT_FIT) = True
       Else
          If phprop.HenrysConstant.CurrentSelection.choice = HENRYS_CONSTANT_FIT Then
             phprop.HenrysConstant.CurrentSelection.choice = 0
             hc_form!lblSource(1).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(HENRYS_CONSTANT_FIT) = False
       End If

       Call DisplayHenrysConstant

      Exit Sub

HenrysConstantUNIFACopTError:
      msg = "Error in the FORTRAN routines while calculating Henry's Constant from UNIFAC at Operating T!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.HenrysConstant.operatingT.UNIFAC.error = -200
      Resume Next

HenrysConstantRegressionError:
      msg = "Error in the FORTRAN routines while calculating Henry's Constant from Regression of Data Points!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.HenrysConstant.regress.error = -200
      Resume Next

VaporPressureErrorHenrysCon:   'Needed to calculate Henry's constants at database Temperatures
      msg = "Error in the FORTRAN routines while calculating Vapor Pressure Needed to Calculate Henry's Constants at Database Temperatures!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      VaporPressureErrHC = -200
      Resume Next

ActivityCoeffErrorHenrysCon:   'Needed to calculate Henry's constants at database Temperatures
      msg = "Error in the FORTRAN routines while calculating Activity Coefficient Needed to Calculate Henry's Constants at Database Temperatures!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      ActivityCoeffErrHC = -200
      Resume Next

HenrysConstantUNIFACdbTError:
      msg = "Error in the FORTRAN routines while calculating Henry's Constant from UNIFAC at Database T!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      HenrysConstantUnifacErr(i) = -200
      Resume Next

HenrysConstantUNIFACfitError:
      msg = "Error in the FORTRAN routines while calculating Henry's Constant from UNIFAC Fit with a Data Point!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.HenrysConstant.fit.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateLiquidDensity()

'       ********************************************************
'       *                                                      *
'       *                Liquid Density                        *
'       *                                                      *
'       ********************************************************

'      /* LDDBCALL:  Get liquid density from the database */
 
      If HaveProperty(MOLECULAR_WEIGHT) Then
         If (phprop.OperatingTemperature > phprop.LiquidDensity.dbase_minT) And (phprop.OperatingTemperature < phprop.LiquidDensity.dbase_maxT) Then
            On Error GoTo LiquidDensityDatabaseError
               Call LDDBCALL(phprop.LiquidDensity.database.Value, phprop.LiquidDensity.database.Source.short, phprop.LiquidDensity.database.Source.long, phprop.LiquidDensity.database.error, phprop.LiquidDensity.database.equation, phprop.LiquidDensity.database.temperature, phprop.LiquidDensity.dbase_minT, phprop.LiquidDensity.dbase_maxT, phprop.LiquidDensity.dbase_coeffA, phprop.LiquidDensity.dbase_coeffB, phprop.LiquidDensity.dbase_coeffC, phprop.LiquidDensity.dbase_coeffD, phprop.MolecularWeight.CurrentSelection.Value, phprop.OperatingTemperature)
         Else 'Temperature is out of range
            phprop.LiquidDensity.database.error = -37
         End If
      Else
         phprop.LiquidDensity.database.error = -43  'Liquid Density can not be calculated in proper units if molecular weight is unavailable
      End If

       If phprop.LiquidDensity.database.error >= 0 Then
          PROPAVAILABLE(LIQUID_DENSITY_DATABASE) = True
       Else
          If phprop.LiquidDensity.CurrentSelection.choice = LIQUID_DENSITY_DATABASE Then
             phprop.LiquidDensity.CurrentSelection.choice = 0
             ldens_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(LIQUID_DENSITY_DATABASE) = False
       End If


'     LDGCCALL:  Obtain Liquid Density from Group Contribution Method

      If HaveProperty(MOLAR_VOLUME_BOILING_POINT) Then
         On Error GoTo LiquidDensityGroupContMethError
            Call LDGCCALL(phprop.LiquidDensity.UNIFAC.Value, phprop.LiquidDensity.UNIFAC.Source.short, phprop.LiquidDensity.UNIFAC.Source.long, phprop.LiquidDensity.UNIFAC.error, phprop.LiquidDensity.UNIFAC.temperature, phprop.MolecularWeight.CurrentSelection.Value, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value, phprop.OperatingTemperature)
      Else
         phprop.LiquidDensity.UNIFAC.error = -13
      End If
      
       If phprop.LiquidDensity.UNIFAC.error >= 0 Then
          PROPAVAILABLE(LIQUID_DENSITY_UNIFAC) = True
       Else
          PROPAVAILABLE(LIQUID_DENSITY_UNIFAC) = False
          If phprop.LiquidDensity.CurrentSelection.choice = LIQUID_DENSITY_UNIFAC Then
             phprop.LiquidDensity.CurrentSelection.choice = 0
             ldens_form!lblSource(1).BackColor = &HC0C0C0
          End If
       End If

      Call DisplayLiquidDensity

      Exit Sub

LiquidDensityDatabaseError:
      msg = "Error in the FORTRAN routines while calculating Liquid Density from the Database!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.LiquidDensity.database.error = -200
      Resume Next

LiquidDensityGroupContMethError:
      msg = "Error in the FORTRAN routines while calculating Liquid Density from Group Contribution Method!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.LiquidDensity.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateLiquidDiffusivity()

'      *******************************************************
'      *                                                     *
'      *              Liquid Diffusivity                     *
'      *                                                     *
'      *******************************************************

'       ********** Calculate liquid diffusivity from
'       ********** Hayduk & Laudie Correlation */

       If HaveProperty(MOLAR_VOLUME_BOILING_POINT) Then

          On Error GoTo LiqDiffHaydukLaudieError
             Call DIFLHL(phprop.LiquidDiffusivity.haydukLaudie.Value, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value, phprop.OperatingTemperature, phprop.MolecularWeight.CurrentSelection.Value, phprop.LiquidDiffusivity.haydukLaudie.error, phprop.LiquidDiffusivity.haydukLaudie.Source.short, phprop.LiquidDiffusivity.haydukLaudie.Source.long, phprop.LiquidDiffusivity.haydukLaudie.temperature)
         
          If phprop.LiquidDiffusivity.haydukLaudie.error >= 0 Then
             PROPAVAILABLE(LIQUID_DIFFUSIVITY_HAYDUKLAUDIE) = True
          Else
             If phprop.LiquidDiffusivity.CurrentSelection.choice = LIQUID_DIFFUSIVITY_HAYDUKLAUDIE Then
                phprop.LiquidDiffusivity.CurrentSelection.choice = 0
                liquid_diff_form!lblSource(0).BackColor = &HC0C0C0
             End If
             PROPAVAILABLE(LIQUID_DIFFUSIVITY_HAYDUKLAUDIE) = False
          End If
       Else
          phprop.LiquidDiffusivity.haydukLaudie.error = -32
          If phprop.LiquidDiffusivity.CurrentSelection.choice = LIQUID_DIFFUSIVITY_HAYDUKLAUDIE Then
             phprop.LiquidDiffusivity.CurrentSelection.choice = 0
             liquid_diff_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(LIQUID_DIFFUSIVITY_HAYDUKLAUDIE) = False
       End If


'     ***** Calculate liquid diffusivity from
'     ***** Method of Polson, 1950

       If HaveProperty(MOLECULAR_WEIGHT) Then
          On Error GoTo LiqDiffPolsonError
             Call DIFLPOL(phprop.LiquidDiffusivity.polson.Value, phprop.MolecularWeight.CurrentSelection.Value, phprop.LiquidDiffusivity.polson.error, phprop.LiquidDiffusivity.polson.Source.short, phprop.LiquidDiffusivity.polson.Source.long, phprop.LiquidDiffusivity.polson.temperature, phprop.OperatingTemperature)
       Else
          phprop.LiquidDiffusivity.polson.error = -47
       End If
       
          If phprop.LiquidDiffusivity.polson.error >= 0 Then
             PROPAVAILABLE(LIQUID_DIFFUSIVITY_POLSON) = True
          Else
             If phprop.LiquidDiffusivity.CurrentSelection.choice = LIQUID_DIFFUSIVITY_POLSON Then
                phprop.LiquidDiffusivity.CurrentSelection.choice = 0
                liquid_diff_form!lblSource(1).BackColor = &HC0C0C0
             End If
             PROPAVAILABLE(LIQUID_DIFFUSIVITY_POLSON) = False
          End If


'     ***** Calculate liquid diffusivity using
'     ***** Wilke-Chang Correlation

       If HaveProperty(MOLAR_VOLUME_BOILING_POINT) Then
          On Error GoTo LiqDiffWilkeChangError
             Call DIFLWC(phprop.LiquidDiffusivity.wilkeChang.Value, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value, phprop.OperatingTemperature, phprop.LiquidDiffusivity.wilkeChang.error, phprop.LiquidDiffusivity.wilkeChang.Source.short, phprop.LiquidDiffusivity.wilkeChang.Source.long, phprop.LiquidDiffusivity.wilkeChang.temperature)

          If phprop.LiquidDiffusivity.wilkeChang.error >= 0 Then
             PROPAVAILABLE(LIQUID_DIFFUSIVITY_WILKECHANG) = True
          Else
             If phprop.LiquidDiffusivity.CurrentSelection.choice = LIQUID_DIFFUSIVITY_WILKECHANG Then
                phprop.LiquidDiffusivity.CurrentSelection.choice = 0
                liquid_diff_form!lblSource(2).BackColor = &HC0C0C0
             End If
             PROPAVAILABLE(LIQUID_DIFFUSIVITY_WILKECHANG) = False
          End If

       Else
          phprop.LiquidDiffusivity.wilkeChang.error = -33
          If phprop.LiquidDiffusivity.CurrentSelection.choice = LIQUID_DIFFUSIVITY_WILKECHANG Then
             phprop.LiquidDiffusivity.CurrentSelection.choice = 0
             liquid_diff_form!lblSource(2).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(LIQUID_DIFFUSIVITY_WILKECHANG) = False
       End If

      Call DisplayLiquidDiffusivity(phprop.MolecularWeight.CurrentSelection.Value)

      Exit Sub

LiqDiffHaydukLaudieError:
      msg = "Error in the FORTRAN routines while calculating Liquid Diffusivity from the Hayduk and Laudie Correlation!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.LiquidDiffusivity.haydukLaudie.error = -200
      Resume Next

LiqDiffPolsonError:
      msg = "Error in the FORTRAN routines while calculating Liquid Diffusivity from the Method of Polson!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.LiquidDiffusivity.polson.error = -200
      Resume Next

LiqDiffWilkeChangError:
      msg = "Error in the FORTRAN routines while calculating Liquid Diffusivity from the Wilke-Chang Correlation!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.LiquidDiffusivity.wilkeChang.error = -200
      Resume Next

End Sub

Sub CalculateMolarVolumeNBP()

'       *********************************************************
'       *                                                       *
'       *     Molar Volume at the Normal Boiling Point          *
'       *                                                       *
'       *********************************************************
      
        If phprop.MaximumUnifacGroups > 0 Then
           On Error GoTo MolarVolumeNBPSchroederError
              Call VBBPCALL(phprop.MolarVolume.BoilingPoint.UNIFAC.Value, phprop.MolarVolume.BoilingPoint.UNIFAC.Source.short, phprop.MolarVolume.BoilingPoint.UNIFAC.Source.long, phprop.MolarVolume.BoilingPoint.UNIFAC.error, phprop.MolarVolume.BoilingPoint.UNIFAC.temperature, phprop.BoilingPoint.database.Value, phprop.MaximumUnifacGroups, phprop.MS(1, 1, 1), phprop.NumberofRingsinCompound)
        Else
           phprop.MolarVolume.BoilingPoint.UNIFAC.error = -36
        End If

       If phprop.MolarVolume.BoilingPoint.UNIFAC.error >= 0 Then
          PROPAVAILABLE(MOLAR_VOLUME_NBP_UNIFAC) = True
       Else
          If phprop.MolarVolume.BoilingPoint.CurrentSelection.choice = MOLAR_VOLUME_NBP_UNIFAC Then
             phprop.MolarVolume.BoilingPoint.CurrentSelection.choice = 0
             mv_nbp_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(MOLAR_VOLUME_NBP_UNIFAC) = False
       End If

     Call DisplayMolarVolumeNBP

      Exit Sub

MolarVolumeNBPSchroederError:
      msg = "Error in the FORTRAN routines while calculating Molar Volume at the Normal Boiling Point from Schroeder's Method!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.MolarVolume.BoilingPoint.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateMolarVolumeOpT()

'        *******************************************************
'        *                                                     *
'        *       Molar Volume at the Operating Temperature     *
'        *                                                     *
'        *******************************************************

'      Calculate Molar Volume at operating temp. from
'      Database liquid density value

       phprop.MolarVolume.operatingT.database.error = 0
       If PROPAVAILABLE(LIQUID_DENSITY_DATABASE) Then
          phprop.MolarVolume.operatingT.database.temperature = phprop.OperatingTemperature
          phprop.MolarVolume.operatingT.database.Source.short = 4

          If HaveProperty(MOLECULAR_WEIGHT) Then
             On Error GoTo MolarVolumeOpTdbError
                Call VBMATT(phprop.MolarVolume.operatingT.database.Value, phprop.LiquidDensity.database.Value, phprop.MolecularWeight.CurrentSelection.Value)
          Else
             phprop.MolarVolume.operatingT.database.error = -55
          End If

       Else
           phprop.MolarVolume.operatingT.database.error = -14
       End If
        
       If phprop.MolarVolume.operatingT.database.error >= 0 Then
          PROPAVAILABLE(MOLAR_VOLUME_OPT_DATABASE) = True
       Else
          If phprop.MolarVolume.operatingT.CurrentSelection.choice = MOLAR_VOLUME_OPT_DATABASE Then
             phprop.MolarVolume.operatingT.CurrentSelection.choice = 0
             molar_vol_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(MOLAR_VOLUME_OPT_DATABASE) = False
       End If


'       /* Calculate Molar Volume at operating temp. from Group Contribution liquid density value */

       phprop.MolarVolume.operatingT.UNIFAC.error = 0
       If PROPAVAILABLE(LIQUID_DENSITY_UNIFAC) Then
          phprop.MolarVolume.operatingT.UNIFAC.temperature = phprop.OperatingTemperature
          phprop.MolarVolume.operatingT.UNIFAC.Source.short = 9

          On Error GoTo MolarVolumeOpTGroupContrError
             Call VBMATT(phprop.MolarVolume.operatingT.UNIFAC.Value, phprop.LiquidDensity.UNIFAC.Value, phprop.MolecularWeight.CurrentSelection.Value)

       Else
          phprop.MolarVolume.operatingT.UNIFAC.error = -15
       End If
        
       If phprop.MolarVolume.operatingT.UNIFAC.error >= 0 Then
          PROPAVAILABLE(MOLAR_VOLUME_OPT_UNIFAC) = True
       Else
          If phprop.MolarVolume.operatingT.CurrentSelection.choice = MOLAR_VOLUME_OPT_UNIFAC Then
             phprop.MolarVolume.operatingT.CurrentSelection.choice = 0
             molar_vol_form!lblSource(1).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(MOLAR_VOLUME_OPT_UNIFAC) = False
       End If

      Call DisplayMolarVolumeOpT

      Exit Sub

MolarVolumeOpTdbError:
      msg = "Error in the FORTRAN routines while calculating Molar Volume at the Operating T Using Database Liquid Density!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.MolarVolume.operatingT.database.error = -200
      Resume Next

MolarVolumeOpTGroupContrError:
      msg = "Error in the FORTRAN routines while calculating Molar Volume at the Operating T Using Liquid Density Value from Group Contribution Method!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.MolarVolume.operatingT.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateMolecularWeight()

'       *********************************************************
'       *                                                       *
'       *               Molecular Weight                        *
'       *                                                       *
'       *********************************************************

'      *** Check if molecular weight available in database

       If phprop.MolecularWeight.database.Value > 0# Then
          PROPAVAILABLE(MOLECULAR_WEIGHT_DATABASE) = True
       Else
          phprop.MolecularWeight.database.error = -42
          If phprop.MolecularWeight.CurrentSelection.choice = MOLAR_WEIGHT_DATABASE Then
             phprop.MolecularWeight.CurrentSelection.choice = 0
             mwt_form!lblSourceLabel(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(MOLECULAR_WEIGHT_DATABASE) = False
       End If


'      *** Calculate Molecular Weight from Group Contribution Method

       If phprop.MaximumUnifacGroups > 0 Then

          On Error GoTo MolecularWtGroupContrError
             Call MWTCALL(phprop.MolecularWeight.UNIFAC.Value, phprop.MolecularWeight.UNIFAC.Source.short, phprop.MolecularWeight.UNIFAC.Source.long, phprop.MolecularWeight.UNIFAC.error, phprop.MaximumUnifacGroups, phprop.MS(1, 1, 1), phprop.XMW(1))
       Else
          phprop.MolecularWeight.UNIFAC.error = -36
       End If

       If phprop.MolecularWeight.UNIFAC.error >= 0 Then
          PROPAVAILABLE(MOLECULAR_WEIGHT_UNIFAC) = True
       Else
          If phprop.MolecularWeight.CurrentSelection.choice = MOLAR_WEIGHT_UNIFAC Then
             phprop.MolecularWeight.CurrentSelection.choice = 0
             mwt_form!lblSourceLabel(1).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(MOLECULAR_WEIGHT_UNIFAC) = False
       End If

       Call DisplayMolecularWeight

      Exit Sub

MolecularWtGroupContrError:
      msg = "Error in the FORTRAN routines while calculating Molecular Weight from Group Contribution Method!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.MolecularWeight.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateOctWaterPartCoeff()

'        ******************************************************
'        *                                                    *
'        *        Octanol Water Partition Coefficient         *
'        *                                                    *
'        ******************************************************


'   /***** VALUE FROM DATABASE */

        If phprop.OctWaterPartCoeff.database.Value = -1 Then
           phprop.OctWaterPartCoeff.database.error = -31
        Else
           phprop.OctWaterPartCoeff.database.error = 0
        End If

       If phprop.OctWaterPartCoeff.database.error >= 0 Then
          PROPAVAILABLE(OCT_WATER_PART_COEFF_DB) = True
       Else
          If phprop.OctWaterPartCoeff.CurrentSelection.choice = OCT_WATER_PART_COEFF_DB Then
             phprop.OctWaterPartCoeff.CurrentSelection.choice = 0
             octanol_form!lblSource(1).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(OCT_WATER_PART_COEFF_DB) = False
       End If


'      /*****  Value from UNIFAC at operating temperature */

      If phprop.MaximumUnifacGroups > 0 Then

CalculateUNIFACKowOperatingT:

          phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value = 0#
          phprop.OctWaterPartCoeff.operatingT.UNIFAC.error = 0

          On Error GoTo OctWatPartCoeffUNIFACopTError
             Call KOWCALL(phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value, phprop.OctWaterPartCoeff.operatingT.UNIFAC.Source.short, phprop.OctWaterPartCoeff.operatingT.UNIFAC.Source.long, phprop.OctWaterPartCoeff.operatingT.UNIFAC.error, phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature, phprop.OperatingTemperature, FGRPErrorFlag, phprop.MaximumUnifacGroups, phprop.MS(1, 1, 1), phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase)

          If phprop.OctWaterPartCoeff.operatingT.UNIFAC.error < 0 Then 'Error calculating solubility with this particular UNIFAC parameter set
             phprop.OctWaterPartCoeff.BinaryInteractionParameterDBAvailable(phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase) = False
             If UserSelectedTheUnifacBIPDBKow Then
                phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = phprop.OctWaterPartCoeff.PreviousBinaryInteractionParameterDB
                MsgBox "Selected UNIFAC database not available to calculate octanol water partition coefficient for this compound.  Returning to Original Choice", MB_ICONSTOP, "Data Not Available"
                octanol_form!cboUNIFACParameterSet.ListIndex = phprop.OctWaterPartCoeff.PreviousBinaryInteractionParameterDB - 1
                GoTo CalculateUNIFACKowOperatingT
             End If

             Select Case phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase
                Case BIP_dbHierarchy.OctWaterPartCoeff(1)
                   phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = BIP_dbHierarchy.OctWaterPartCoeff(2)
                   GoTo CalculateUNIFACKowOperatingT
                Case BIP_dbHierarchy.OctWaterPartCoeff(2)
                    phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = 0
             End Select
          End If
          If phprop.OctWaterPartCoeff.operatingT.UNIFAC.error < 0 Then
             phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = 0
          End If
       Else
          phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = 0
          phprop.OctWaterPartCoeff.operatingT.UNIFAC.error = -36
       End If

       If phprop.OctWaterPartCoeff.operatingT.UNIFAC.error >= 0 Then
          PROPAVAILABLE(OCT_WATER_PART_COEFF_OPT_UNIFAC) = True
       Else
          If phprop.OctWaterPartCoeff.CurrentSelection.choice = OCT_WATER_PART_COEFF_OPT_UNIFAC Then
             phprop.OctWaterPartCoeff.CurrentSelection.choice = 0
             octanol_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(OCT_WATER_PART_COEFF_OPT_UNIFAC) = False
       End If


'     *****  Value from UNIFAC at database temperature

       If PROPAVAILABLE(OCT_WATER_PART_COEFF_DB) Then
          phprop.OctWaterPartCoeff.databaseT.UNIFAC.error = 0
          If phprop.MaximumUnifacGroups > 0 Then
             On Error GoTo OctWatPartCoeffdbTError
                Call KOWCALL(phprop.OctWaterPartCoeff.databaseT.UNIFAC.Value, phprop.OctWaterPartCoeff.databaseT.UNIFAC.Source.short, phprop.OctWaterPartCoeff.databaseT.UNIFAC.Source.long, phprop.OctWaterPartCoeff.databaseT.UNIFAC.error, phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature, phprop.OctWaterPartCoeff.database.temperature, FGRPErrorFlag, phprop.MaximumUnifacGroups, phprop.MS(1, 1, 1), phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase)
          Else
             phprop.OctWaterPartCoeff.databaseT.UNIFAC.error = -36
          End If
       Else
          phprop.OctWaterPartCoeff.databaseT.UNIFAC.error = -46
       End If


       If phprop.OctWaterPartCoeff.databaseT.UNIFAC.error >= 0 Then
          PROPAVAILABLE(OCT_WATER_PART_COEFF_DBT_UNIFAC) = True
       Else
          If phprop.OctWaterPartCoeff.CurrentSelection.choice = OCT_WATER_PART_COEFF_DBT_UNIFAC Then
             phprop.OctWaterPartCoeff.CurrentSelection.choice = 0
             octanol_form!lblSource(2).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(OCT_WATER_PART_COEFF_DBT_UNIFAC) = False
       End If

      Call DisplayOctWaterPartCoeff

      Exit Sub

OctWatPartCoeffUNIFACopTError:
      msg = "Error in the FORTRAN routines while calculating Octanol Water Partition Coefficient from UNIFAC at Operating T!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.OctWaterPartCoeff.operatingT.UNIFAC.error = -200
      Resume Next

OctWatPartCoeffdbTError:
      msg = "Error in the FORTRAN routines while calculating Octanol Water Partition Coefficient from UNIFAC at Database T!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.OctWaterPartCoeff.databaseT.UNIFAC.error = -200
      Resume Next

End Sub

Sub CalculateProperties()
    Dim i As Integer, msg As String

    Call CalculateVaporPressure
    contam_prop_form.Refresh

    Call CalculateActivityCoefficient
    contam_prop_form.Refresh

    Call CalculateHenrysConstant
    contam_prop_form.Refresh

    Call CalculateMolecularWeight
    contam_prop_form.Refresh

    Call CalculateBoilingPoint
    contam_prop_form.Refresh

    Call CalculateMolarVolumeNBP
    contam_prop_form.Refresh

    Call CalculateLiquidDensity
    contam_prop_form.Refresh

    Call CalculateMolarVolumeOpT
    contam_prop_form.Refresh

    Call CalculateRefractiveIndex
    contam_prop_form.Refresh

    Call CalculateAqueousSolubility
    contam_prop_form.Refresh

    Call CalculateOctWaterPartCoeff
    contam_prop_form.Refresh
        
    Call CalculateLiquidDiffusivity
    contam_prop_form.Refresh

    Call CalculateGasDiffusivity
    contam_prop_form.Refresh

    Call CalculateWaterDensity
    contam_prop_form.Refresh

    Call CalculateWaterViscosity
    contam_prop_form.Refresh

    Call CalculateWaterSurfaceTension
    contam_prop_form.Refresh

    Call CalculateAirDensity
    contam_prop_form.Refresh

    Call CalculateAirViscosity
    contam_prop_form.Refresh


'*** Place PROPAVAILABLE and HAVEPROPERTY arrays into phprop structure
     For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
         phprop.PROPAVAILABLE(i) = PROPAVAILABLE(i)
     Next i
     For i = 1 To NUMBER_OF_PROPERTIES
         phprop.HaveProperty(i) = HaveProperty(i)
     Next i

End Sub

Sub CalculateRefractiveIndex()

'        ******************************************************
'        *                                                    *
'        *               Refractive Index                     *
'        *                                                    *
'        ******************************************************

'        /* Note:  This is only available from the database so
'                  no UNIFAC calculations are needed for it */
      
      If phprop.RefractiveIndex.database.Value = -1 Then
         phprop.RefractiveIndex.database.error = -17
      Else
         phprop.RefractiveIndex.database.error = 0
      End If

       If phprop.RefractiveIndex.database.error >= 0 Then
          PROPAVAILABLE(REFRACTIVE_INDEX_DATABASE) = True
       Else
          If phprop.RefractiveIndex.CurrentSelection.choice = REFRACTIVE_INDEX_DATABASE Then
             phprop.RefractiveIndex.CurrentSelection.choice = 0
             rindex_form!lblSource(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(REFRACTIVE_INDEX_DATABASE) = False
       End If

      Call DisplayRefractiveIndex

End Sub

Sub CalculateVaporPressure()

'       *********************************************************
'       *                                                       *
'       *                  Vapor Pressure                       *
'       *                                                       *
'       *********************************************************
     
       On Error GoTo VaporPressureDatabaseError
          Call VPRCALL( _
              phprop.VaporPressure.database.Value, _
              phprop.VaporPressure.database.Source.short, _
              phprop.VaporPressure.database.Source.long, _
              phprop.VaporPressure.database.error, _
              phprop.VaporPressure.database.equation, _
              phprop.VaporPressure.database.temperature, _
              phprop.VaporPressure.database.minimumT, _
              phprop.VaporPressure.database.maximumT, _
              phprop.VaporPressure.database.antoineA, _
              phprop.VaporPressure.database.antoineB, _
              phprop.VaporPressure.database.antoineC, _
              phprop.VaporPressure.database.antoineD, _
              phprop.VaporPressure.database.antoineE, _
              phprop.VaporPressure.database.superfund.Value, _
              phprop.VaporPressure.database.superfund.temperature, _
              phprop.OperatingTemperature)

       If phprop.VaporPressure.database.error >= 0 Then
          PROPAVAILABLE(VAPOR_PRESSURE_DATABASE) = True
       Else
          If phprop.VaporPressure.CurrentSelection.choice = VAPOR_PRESSURE_DATABASE Then
             phprop.VaporPressure.CurrentSelection.choice = 0
             vp_form!lblSourceLabel(0).BackColor = &HC0C0C0
          End If
          PROPAVAILABLE(VAPOR_PRESSURE_DATABASE) = False
       End If

      Call DisplayVaporPressure

      Exit Sub

VaporPressureDatabaseError:
      msg = "Error in the FORTRAN routines while calculating Vapor Pressure from the Database!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.VaporPressure.database.error = -200
      Resume Next

End Sub

Sub CalculateWaterDensity()

'       ********************************************************
'       *                                                      *
'       *              Water Density                           *
'       *                                                      *
'       ********************************************************

      On Error GoTo WaterDensityCorrelationError
         Call H2ODENS(phprop.WaterDensity.correlation.Value, phprop.OperatingTemperature, phprop.WaterDensity.correlation.error, phprop.WaterDensity.correlation.Source.short, phprop.WaterDensity.correlation.Source.long, phprop.WaterDensity.correlation.temperature)

      If (phprop.OperatingTemperature < 0#) Or (phprop.OperatingTemperature > 100#) Then 'Temperature is out of the valid range for the water density correlation
         phprop.WaterDensity.correlation.error = 11
      End If

      If phprop.WaterDensity.correlation.error >= 0 Then
         PROPAVAILABLE(WATER_DENSITY_CORRELATION) = True
      Else
         If phprop.WaterDensity.CurrentSelection.choice = WATER_DENSITY_CORRELATION Then
            phprop.WaterDensity.CurrentSelection.choice = 0
            frmWaterDensity!lblSource(0).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(WATER_DENSITY_CORRELATION) = False
      End If

       Call DisplayWaterDensity

      Exit Sub

WaterDensityCorrelationError:
      msg = "Error in the FORTRAN routines while calculating Water Density from Correlation!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.WaterDensity.correlation.error = -200
      Resume Next

End Sub

Sub CalculateWaterSurfaceTension()

'      ********************************************************
'      *                                                      *
'      *              Water Surface Tension                   *
'      *                                                      *
'      ********************************************************

      On Error GoTo WaterSurfTensCorrelationError
         Call H2OST(phprop.WaterSurfaceTension.correlation.Value, phprop.OperatingTemperature, phprop.WaterSurfaceTension.correlation.error, phprop.WaterSurfaceTension.correlation.Source.short, phprop.WaterSurfaceTension.correlation.Source.long, phprop.WaterSurfaceTension.correlation.temperature)

      If phprop.WaterSurfaceTension.correlation.error >= 0 Then
         PROPAVAILABLE(WATER_SURF_TENSION_CORRELATION) = True
      Else
         If phprop.WaterSurfaceTension.CurrentSelection.choice = WATER_SURF_TENSION_CORRELATION Then
            phprop.WaterSurfaceTension.CurrentSelection.choice = 0
            frmWaterSurfaceTension!lblSource(0).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(WATER_SURF_TENSION_CORRELATION) = False
      End If

      Call DisplayWaterSurfaceTension

      Exit Sub

WaterSurfTensCorrelationError:
      msg = "Error in the FORTRAN routines while calculating Water Surface Tension from Correlation!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.WaterSurfaceTension.correlation.error = -200
      Resume Next

End Sub

Sub CalculateWaterViscosity()

'       ********************************************************
'       *                                                      *
'       *              Water Viscosity                         *
'       *                                                      *
'       ********************************************************

      On Error GoTo WaterViscosityCorrelationError
         Call H2OVISC(phprop.WaterViscosity.correlation.Value, phprop.OperatingTemperature, phprop.WaterViscosity.correlation.error, phprop.WaterViscosity.correlation.Source.short, phprop.WaterViscosity.correlation.Source.long, phprop.WaterViscosity.correlation.temperature)

      If phprop.WaterViscosity.correlation.error >= 0 Then
         PROPAVAILABLE(WATER_VISCOSITY_CORRELATION) = True
      Else
         If phprop.WaterViscosity.CurrentSelection.choice = WATER_VISCOSITY_CORRELATION Then
            phprop.WaterViscosity.CurrentSelection.choice = 0
            frmWaterViscosity!lblSource(0).BackColor = &HC0C0C0
         End If
         PROPAVAILABLE(WATER_VISCOSITY_CORRELATION) = False
      End If

       Call DisplayWaterViscosity

      Exit Sub

WaterViscosityCorrelationError:
      msg = "Error in the FORTRAN routines while calculating Water Viscosity from Correlation!"
      MsgBox msg, MB_ICONINFORMATION, "Error"
      phprop.WaterViscosity.correlation.error = -200
      Resume Next

End Sub

Sub DoCalculationForThisContaminant()

    Static HaveValue(NUMBER_OF_PROPERTIES_AVAILABLE)     As Long '/* array to store whether we have a value for a particular physical property, either from the database, from UNIFAC, or from user-input */
    Dim i As Long
 '   Dim FGRPErrorFlag As Long ' /* Variable to store if there is an error in retrieving the UNIFAC functional groups */
    Static HCDatabaseTemperature(NDCONSTANT) As Double
    Static HCUnifacValue(NDCONSTANT) As Double
    Static HCUnifacError(NDCONSTANT) As Long
    Static HCunifacSourceShorter(NDCONSTANT) As Long
    Static HCunifacSourceLonger(NDCONSTANT) As Long
    Dim MolecularWeight As Double '/* Variable for use in call to LDDBCALL */
    Dim hc_database_value As String * 40
    Dim hc_database_temp As String
    Dim hc_string As String
    Dim J As Integer, K As Integer, L As Integer


    '   /************************************************************
    '    *                                                          *
    '    *   Set input variables equal to appropriate records in    *
    '    *        the physical properties structure                 *
    '    *                                                          *
    '    ************************************************************/

    phprop.CasNumber = dbinput.CasNumber
    phprop.Name = dbinput.Name
    phprop.formula = dbinput.formula

    phprop.MolecularWeight.database.Value = dbinput.MolecularWeight
    phprop.MolecularWeight.database.Source.short = 4  '/* Source for molecular weight = DIPPR801 */

    For i = 1 To dbinput.NumberOfDatabaseHenrysConstants
'        Call GET_HRY(I, phprop.HenrysConstant.database(I).Value)
'        Call GET_TMP1(I, phprop.HenrysConstant.database(I).Temperature)
        phprop.HenrysConstant.database(i).Value = dbinput.HenrysConstant(i)
        phprop.HenrysConstant.database(i).temperature = dbinput.HenrysConstantTemperature(i)
        phprop.HenrysConstant.database(i).Source.short = dbinput.HenrysConstantSource
        phprop.HenrysConstant.database(i).error = 0
    Next i

    phprop.HenrysConstant.NumberOfDatabaseHenrysConstants = dbinput.NumberOfDatabaseHenrysConstants

    phprop.VaporPressure.database.superfund.Value = dbinput.VaporPressureSuperfund
    phprop.VaporPressure.database.superfund.temperature = dbinput.VaporPressureSuperfundTemperature

    phprop.LiquidDensity.database.equation = dbinput.LiquidDensityEquation
    phprop.LiquidDensity.dbase_n_coeffs = dbinput.LiquidDensityNumberCoefficients
    phprop.LiquidDensity.dbase_coeffA = dbinput.LiquidDensityCoefficientA
    phprop.LiquidDensity.dbase_coeffB = dbinput.LiquidDensityCoefficientB
    phprop.LiquidDensity.dbase_coeffC = dbinput.LiquidDensityCoefficientC
    phprop.LiquidDensity.dbase_coeffD = dbinput.LiquidDensityCoefficientD
    phprop.LiquidDensity.dbase_minT = dbinput.LiquidDensityMinimumT - 273.15
    phprop.LiquidDensity.dbase_maxT = dbinput.LiquidDensityMaximumT - 273.15
    phprop.LiquidDensity.database.Source.short = dbinput.LiquidDensitySource

    phprop.VaporPressure.database.equation = dbinput.VaporPressureDatabaseEquation
    phprop.VaporPressure.database.ncoeffs = dbinput.VaporPressureNumberCoefficients
    phprop.VaporPressure.database.antoineA = dbinput.VaporPressureAntoineA
    phprop.VaporPressure.database.antoineB = dbinput.VaporPressureAntoineB
    phprop.VaporPressure.database.antoineC = dbinput.VaporPressureAntoineC
    phprop.VaporPressure.database.antoineD = dbinput.VaporPressureAntoineD
    phprop.VaporPressure.database.antoineE = dbinput.VaporPressureAntoineE
    phprop.VaporPressure.database.minimumT = dbinput.VaporPressureMinimumT - 273.15
    phprop.VaporPressure.database.maximumT = dbinput.VaporPressureMaximumT - 273.15
    phprop.VaporPressure.database.Source.short = dbinput.VaporPressureSource

    phprop.AqueousSolubility.database.Value = dbinput.AqueousSolubility
    phprop.AqueousSolubility.database.temperature = dbinput.AqueousSolubilityTemperature
    phprop.AqueousSolubility.database.Source.short = dbinput.AqueousSolubilitySource

    phprop.OctWaterPartCoeff.database.Value = dbinput.OctWaterPartCoeff
    phprop.OctWaterPartCoeff.database.temperature = dbinput.OctWaterPartCoeffTemperature
    phprop.OctWaterPartCoeff.database.Source.short = dbinput.OctWaterPartCoeffSource

    phprop.BoilingPoint.database.Value = dbinput.BoilingPoint - 273.15
    phprop.BoilingPoint.database.Source.short = dbinput.BoilingPointSource

    phprop.RefractiveIndex.database.Value = dbinput.RefractiveIndex
    phprop.RefractiveIndex.database.Source.short = dbinput.RefractiveIndexSource

    phprop.OperatingTemperature = dbinput.OperatingTemperature

    For J = 1 To 10
        For K = 1 To 10
            For L = 1 To 2
                phprop.MS(J, K, L) = dbinput.MS(J, K, L)
            Next L
        Next K
    Next J
    phprop.MaximumUnifacGroups = dbinput.MaximumUnifacGroups
    phprop.NumberofRingsinCompound = dbinput.NumberofRingsinCompound

    Call CalculateProperties
    PropContaminant(NumSelectedChemicals) = phprop
    PreviouslySelectedIndex = NumSelectedChemicals

End Sub

Sub GetClosestHCDatabaseT()
    Dim CurrDiff As Double
    Dim PermDiff As Double
    Dim CurrIndex As Integer
    Dim PermIndex As Integer

    PermDiff = Abs(phprop.OperatingTemperature - phprop.HenrysConstant.database(1).temperature)
    PermIndex = 1
    For CurrIndex = 2 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
        CurrDiff = Abs(phprop.OperatingTemperature - phprop.HenrysConstant.database(CurrIndex).temperature)
        If CurrDiff < PermDiff Then
           PermDiff = CurrDiff
           PermIndex = CurrIndex
        End If
    Next CurrIndex
    phprop.HenrysConstant.chosenDatabaseIndex = PermIndex

End Sub

Sub GetClosestHCUnifacT()
    Dim CurrDiff As Double
    Dim PermDiff As Double
    Dim CurrIndex As Integer
    Dim PermIndex As Integer

    PermDiff = 1E+40
    PermIndex = 0
    For CurrIndex = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
        If phprop.HenrysConstant.UNIFAC(CurrIndex).error >= 0 Then
           CurrDiff = Abs(phprop.OperatingTemperature - phprop.HenrysConstant.UNIFAC(CurrIndex).temperature)
           If CurrDiff < PermDiff Then
              PermDiff = CurrDiff
              PermIndex = CurrIndex
           End If
        End If
    Next CurrIndex
    phprop.HenrysConstant.chosenUNIFACIndex = PermIndex

End Sub

