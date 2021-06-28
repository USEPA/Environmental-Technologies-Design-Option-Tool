Option Strict Off
Option Explicit On
Module Validations
	
	Sub Check_Length(ByRef D As Double, ByRef Flag_Small As Short)
		Dim E As Double
		Dim msg As String
		E = Bed.Weight * 4# / PI / Bed.Diameter ^ 2 / D / Carbon.Density / 1000#
		If (E > 0.999) Then
			msg = "The bed length you just entered is too small to contain the weight of activated carbon you have specified."
			msg = msg & Chr(10) & Chr(10) & "Recommendation: Try to re-specify bed diameter, bed length, weight of carbon, and/or carbon density."
			Call Show_Error(msg)
			Flag_Small = True
		Else
			Flag_Small = False
		End If
	End Sub
	Sub Check_Density(ByRef D As Double, ByRef Flag_Small As Short)
		Dim E As Double
		Dim msg As String
		E = Bed.Weight * 4# / PI / Bed.Diameter ^ 2 / Bed.length / D / 1000#
		If E > 0.999 Then
			msg = "The apparent density of carbon you just entered creates a bed dimensioning scenario where the bed is too small to contain the weight of activated carbon you have specified."
			msg = msg & Chr(10) & Chr(10) & "Recommendation: Try to re-specify bed diameter, bed length, weight of carbon, and/or carbon density."
			Call Show_Error(msg)
			Flag_Small = True
		Else
			Flag_Small = False
		End If
	End Sub
	Sub Check_Diameter(ByRef D As Double, ByRef Flag_Small As Short)
		Dim E As Double
		Dim msg As String
		E = Bed.Weight * 4# / PI / D ^ 2 / Bed.length / Carbon.Density / 1000#
		If E > 0.999 Then
			msg = "The bed diameter you just entered is too small to contain the weight of activated carbon you have specified."
			msg = msg & Chr(10) & Chr(10) & "Recommendation: Try to re-specify bed diameter, bed length, weight of carbon, and/or carbon density."
			Call Show_Error(msg)
			Flag_Small = True
		Else
			Flag_Small = False
		End If
	End Sub
	Sub Check_Weight(ByRef W As Double, ByRef Flag_Small As Short)
		Dim E As Double
		Dim msg As String
		E = W * 4# / PI / Bed.Diameter ^ 2 / Bed.length / Carbon.Density / 1000#
		If E > 0.999 Then
			msg = "The weight of activated carbon you just entered is too large to be contained by the bed with your specified dimensions."
			msg = msg & Chr(10) & Chr(10) & "Recommendation: Try to re-specify bed diameter, bed length, weight of carbon, and/or carbon density."
			Call Show_Error(msg)
			Flag_Small = True
		Else
			Flag_Small = False
		End If
	End Sub
	
	
	Sub Check_Time_Parameters(ByRef WhichParam As Short, ByRef EnteredValue As Double, ByRef ForceAbort As Short)
		Dim New_FirstPoint As Double
		Dim New_LastPoint As Double
		Dim New_TimeStep As Double
		Dim UserMsg As String
		Dim Force_New_TimeStep As Double
		Dim Old_TimeStep As Double
		Dim conversionfactor1 As Double
		Dim conversionfactor2 As Double
		
		ForceAbort = False
		
		New_FirstPoint = TimeP.Init
		New_LastPoint = TimeP.End_Renamed
		New_TimeStep = TimeP.Step_Renamed
		Select Case WhichParam
			Case 0 'Total run time
				New_LastPoint = EnteredValue
			Case 1 'First point
				New_FirstPoint = EnteredValue
			Case 2 'Time step
				New_TimeStep = EnteredValue
		End Select
		
		If (New_FirstPoint > New_LastPoint) Then
			ForceAbort = True
			Call Show_Error("The first time cannot be greater than the " & "final time.")
			Exit Sub
		End If
		'If (New_TimeStep < ((New_LastPoint - New_FirstPoint) / (Number_Points_Max - 1))) Then
		If (New_FirstPoint < New_LastPoint) Then
			'upon legit times entered, execute to adjust time step

			'set timestep to max value
			'New_TimeStep = (New_LastPoint - New_FirstPoint) / (Number_Points_Max - 1)
			ForceAbort = True
			'Call Show_Error("The time step is too small.  The maximum " & "number of points is " & Trim(Str(Number_Points_Max)) & ".")
			Old_TimeStep = New_TimeStep
			Force_New_TimeStep = (New_LastPoint - New_FirstPoint) / (Number_Points_Max - 1)
			'If (New_TimeStep < Force_New_TimeStep) Then
			New_TimeStep = Force_New_TimeStep
				TimeP.Step_Renamed = New_TimeStep
				TimeP.Init = New_FirstPoint
				TimeP.End_Renamed = New_LastPoint
				UserMsg = UserMsg & "  In addition, the time step was adjusted from " & Format_It(Old_TimeStep, 3) & " minutes to " & Format_It(New_TimeStep, 3) & " minutes."
			'End If
			'Call Show_Message(UserMsg)
			Exit Sub
		End If
		If ((Bed.NumberOfBeds = 1) Or (New_FirstPoint < 0.00011)) Then
			'Do nothing--this is okay.
		Else
			'For beds in series, initial time must be approximately zero.
			ForceAbort = True
			Old_TimeStep = New_TimeStep
			UserMsg = "For more than one axial element, the initial time must be approximately zero.  The initial time has been automatically reset to 0.0001 minutes."
			New_FirstPoint = 0.0001
			Force_New_TimeStep = (New_LastPoint - New_FirstPoint) / (Number_Points_Max - 5)
			If (New_TimeStep < Force_New_TimeStep) Then
				New_TimeStep = Force_New_TimeStep
				UserMsg = UserMsg & "  In addition, the time step was adjusted from " & Format_It(Old_TimeStep, 3) & " minutes to " & Format_It(New_TimeStep, 3) & " minutes."
			End If
			Call Show_Error(UserMsg)
			
			'THIS CODE WAS INTENDED TO REDISPLAY TIMES TO MAIN WINDOW.
			'IT IS NO LONGER NEEDED BECAUSE A CALL TO frmMain_Refresh() OCCURS
			'RIGHT AFTER THE CALL TO THIS SUBROUTINE.
			'conversionfactor1 = TimeConversionFactor(CInt(frmMain.txttimeunits(1).ListIndex))
			'conversionfactor2 = TimeConversionFactor(CInt(frmMain.txttimeunits(2).ListIndex))
			'txtTime(1) = Format_It(New_FirstPoint * 60 * conversionfactor1, 3)
			'txtTime(2) = Format_It(New_TimeStep * 60 * conversionfactor2, 3)
			TimeP.Init = New_FirstPoint
			TimeP.Step_Renamed = New_TimeStep
			
			Exit Sub
		End If
		
		
		
		'If FirstPT > EndT Then
		'  MsgBox "The first point is greater than the final point.", MB_ICONEXCLAMATION, AppName_For_Display_long
		'  Exit Sub
		'ElseIf Time_Step < ((EndT - FirstPT) / (Number_Points_Max - 1)) Then
		'  MsgBox "Time step is too small. The maximum number of points is 400.", MB_ICONEXCLAMATION, AppName_For_Display_long
		'  Exit Sub
		'End If
		'TimeP.End = EndT * 24 * 60#    'To convert from days to minutes
		'If (Bed.NumberOfBeds = 1) Or ((FirstPT * 24# * 60#) < .00011) Then
		'   TimeP.Init = FirstPT * 24# * 60#  'To convert from days to minutes
		'   TimeP.Step = Time_Step * 24# * 60# 'To convert from days to minutes
		'Else   'For beds in series, initial time must be approximately zero
		'   FirstPT = .0001 / 24# / 60#
		'   NewTimeStep = (EndT - FirstPT) / (Number_Points_Max - 5)
		'   If Time_Step < NewTimeStep Then Time_Step = NewTimeStep
		'   MsgBox "For beds in series, the initial time must be approximately zero.  The initial time will automatically be adjusted to reflect this.  If necessary, the time step will also be adjusted.", MB_ICONINFORMATION
		'   txtTime(1) = Format_It(FirstPT, 2)
		'   txtTime(2) = Format_It(Time_Step, 2)
		'   Exit Sub
		'End If
		
	End Sub
End Module