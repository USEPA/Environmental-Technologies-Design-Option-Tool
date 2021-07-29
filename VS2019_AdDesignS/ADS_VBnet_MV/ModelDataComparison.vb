Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports System.Windows.Forms.DataVisualization.Charting
Friend Class frmModelDataComparison
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer

	Dim Cin() As Double
	Dim Td() As Double
	Dim Cd() As Double
	Dim Fmin() As Double
	Dim UnloadMe As Short
	
	Dim PopulatingScrollboxes As Short
	
	'---- Conc Units
	Const CBOCUNITS_CC0 As Short = 0
	Const CBOCUNITS_mg_L As Short = 1
	Const CBOCUNITS_ug_L As Short = 2
	
	'---- Time Units
	Const CBOTUNITS_days As Short = 0
	Const CBOTUNITS_BVF As Short = 1
	Const CBOTUNITS_VTM As Short = 2
	
	
	
	
	
	Const frmModelDataComparison_declarations_end As Boolean = True
	
	
	Private Sub Populate_Scrollboxes()
		Dim i As Short
		PopulatingScrollboxes = True
		'For i = 0 To 2
		'  cboDataset(i).Clear
		'  cboDataset(i).AddItem "Off"
		'  cboDataset(i).AddItem "Symbols"
		'  cboDataset(i).AddItem "Lines"
		'  cboDataset(i).AddItem "Symbols and Lines"
		'Next i
		cboCUnits.Items.Clear()
		cboCUnits.Items.Add("C/C0")
		cboCUnits.Items.Add("mg/L")
		cboCUnits.Items.Add(Chr(181) & "g/L")
		cboTUnits.Items.Clear()
		cboTUnits.Items.Add("days")
		cboTUnits.Items.Add("BVF")
		cboTUnits.Items.Add("VTM")
		cboGraphType.Items.Clear()
		cboGraphType.Items.Add("Dots")
		cboGraphType.Items.Add("Lines")
		cboGrid.Items.Add("None")
		cboGrid.Items.Add("Horizontal")
		cboGrid.Items.Add("Vertical")
		cboGrid.Items.Add("Both")
		cboCompo.Items.Clear()
		Select Case frmCompareData_WhichSet
			Case frmCompareData_WhichSet_PSDM
				For i = 1 To Results.NComponent
					cboCompo.Items.Add(Trim(Results.Component(i).Name))
				Next i
			Case frmCompareData_WhichSet_CPHSDM
				cboCompo.Items.Add(Trim(CPM_Results.Component.Name))
		End Select
		'---- Read in INI settings
		'cboDataset(0).ListIndex = 1
		'cboDataset(1).ListIndex = 1
		'cboDataset(2).ListIndex = 1
		cboCUnits.SelectedIndex = 0
		cboTUnits.SelectedIndex = 0
		cboGraphType.SelectedIndex = 0
		cboGrid.SelectedIndex = 0
		cboCompo.SelectedIndex = 0
		Call UserPrefs_Load()
		PopulatingScrollboxes = False
	End Sub
	
	
	
	
	'UPGRADE_WARNING: Event cboCompo.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCompo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCompo.SelectedIndexChanged
		If (Not PopulatingScrollboxes) Then
			Call Draw_Curves(cboCompo.SelectedIndex + 1)
			'lblErrorC = Format$(Fmin(cboCompo.ListIndex + 1), "0.0000E+00")
		End If
	End Sub
	'UPGRADE_WARNING: Event cboCUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCUnits.SelectedIndexChanged
		If (Not PopulatingScrollboxes) Then
			Call Draw_Curves(cboCompo.SelectedIndex + 1)
		End If
	End Sub
	Private Sub cboDataset_Click(ByRef Index As Short)
		'If (Not PopulatingScrollboxes) Then
		'  Call Draw_Curves(cboCompo.ListIndex + 1)
		'End If
	End Sub
	'UPGRADE_WARNING: Event cboGraphType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboGraphType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGraphType.SelectedIndexChanged
		If (Not PopulatingScrollboxes) Then
			Select Case cboGraphType.SelectedIndex
				Case 0 'Dots
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Chart1.Series(1).ChartType = SeriesChartType.Point
					'Chart1.Series(2).ChartType = SeriesChartType.Point
					'grpBreak.GraphStyle = 1
				Case 1 'Lines
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Chart1.Series(1).ChartType = SeriesChartType.Line
					'Chart1.Series(2).ChartType = SeriesChartType.Line
					'grpBreak.GraphStyle = 4
			End Select
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.DrawMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.DrawMode = 2
		End If
	End Sub
	'UPGRADE_WARNING: Event cboGrid.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboGrid_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGrid.SelectedIndexChanged
		Dim gridstyle As Integer
		If (Not PopulatingScrollboxes) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GridStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gridstyle = cboGrid.SelectedIndex + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.DrawMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case (gridstyle)
				Case 1
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 0
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 0
				Case 2
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 0
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 1
				Case 3
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 1
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 0
				Case 4
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 1
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 1
			End Select

		End If
	End Sub
	'UPGRADE_WARNING: Event cboTUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboTUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboTUnits.SelectedIndexChanged
		If (Not PopulatingScrollboxes) Then
			Call Draw_Curves(cboCompo.SelectedIndex + 1)
		End If
	End Sub
	
	
	Private Sub cmdClose_Click()
		Me.Close()
	End Sub
	
	
	Private Sub cmdPrint_Click()
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()

		'	'''Dim i As Integer
		'		'''Dim H As Single
		'		'''Dim W As Single
		'		'''  'printer.ScaleLeft = -1080  'Set a 3/4-inch margin
		'		'''  'printer.ScaleTop = -1080
		'		'''  'printer.CurrentX = 0
		'		'''  'printer.CurrentY = 0
		'		'''  '
		'		'''  '  printer.FontSize = 12
		'		'''  '  printer.FontBold = True
		'		'''  '  printer.FontUnderline = True
		'		'''  '  printer.Print "Input data for the Plug-Flow Pore And Surface Diffusion Model"
		'		'''  '  printer.FontSize = 10
		'		'''  '  printer.FontBold = False
		'		'''  '  printer.FontUnderline = False
		'		'''  '  '-- Print Filename
		'		'''  '  printer.Print
		'		'''  '  printer.Print "From Data File: "; Filename
		'		'''  '---- Print the graph ------------------------
		'		'''  For i = 1 To Number_Component
		''		'''    grpBreak.ThisPoint = i
		'	'''    grpBreak.PatternData = i - 1
		'	'''  Next i
		'	'''  H = grpBreak.Height
		'	'''  W = grpBreak.Width
		'	'''  grpBreak.Visible = False 'Hide it before printing
		'	'''  If (Printer.Width < Printer.Height) Then
		'	'''    grpBreak.Height = CSng(Printer.Height / 2#)
		'	'''    grpBreak.Width = Printer.Width
		'	'''  Else
		'	'''    grpBreak.Height = Printer.Height
		'	'''    grpBreak.Width = Printer.Width
		'	'''  End If
		'	'''  grpBreak.PrintStyle = 2
		'	'''  grpBreak.DrawMode = 5
		'	'''  grpBreak.Height = H
		'	'''  grpBreak.Width = W
		'	'''  grpBreak.Visible = True
		'	'''  grpBreak.PrintStyle = 2
		'	'''  grpBreak.DrawMode = 2
		'	'''  Printer.EndDoc

	End Sub
	
	
	Private Sub Draw_Curves(ByRef Component_Index As Short)
		Dim i, J As Short
		Dim Data_Max, t_factor As Double
		Dim Bottom_Title As String
		Dim c_factor As Double
		Dim Left_Title As String
		Dim bigger As Short
		Dim SameX As Object
		Dim SameY As Double
		Dim LastPointI As Short
		Dim bed_data As BedPropertyType
		'UPGRADE_WARNING: Arrays in structure comp_data may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim comp_data As ComponentPropertyType
		Dim num_model_points As Short

		Dim numsets, numpoints(2) As Short
		Dim myX, myY As Double

		Chart1.Series.Clear()
		Chart1.ChartAreas(0).RecalculateAxesScale()

		'	grpBreak.DrawMode = 1   empty graph
		Select Case frmCompareData_WhichSet
			Case frmCompareData_WhichSet_PSDM
				'UPGRADE_WARNING: Couldn't resolve default property of object bed_data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				bed_data = Results.Bed
				'UPGRADE_WARNING: Couldn't resolve default property of object comp_data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				comp_data = Results.Component(Component_Index)
				num_model_points = Results.npoints
			Case frmCompareData_WhichSet_CPHSDM
				'UPGRADE_WARNING: Couldn't resolve default property of object bed_data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				bed_data = CPM_Results.Bed
				'UPGRADE_WARNING: Couldn't resolve default property of object comp_data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				comp_data = CPM_Results.Component
				num_model_points = 100
		End Select
		
		Select Case cboTUnits.SelectedIndex
			Case CBOTUNITS_days
				t_factor = 1# / 60# / 24# 'mn -> days
				Bottom_Title = "Time(days)"
			Case CBOTUNITS_BVF
				t_factor = (60.0# * bed_data.Flowrate / bed_data.length / PI / (bed_data.Diameter / 2.0#) ^ 2) / 1000
				Bottom_Title = "Bed Volumes Treated (Thousands)"
			Case CBOTUNITS_VTM
				t_factor = 60# * bed_data.Flowrate / bed_data.Weight
				Bottom_Title = "m" & Chr(179) & " treated per kg of adsorbent"
		End Select
		
		Select Case cboCUnits.SelectedIndex
			Case CBOCUNITS_CC0
				c_factor = 1#
				Left_Title = "C/C0"
			Case CBOCUNITS_mg_L
				c_factor = comp_data.InitialConcentration
				Left_Title = "mg/L"
			Case CBOCUNITS_ug_L
				c_factor = comp_data.InitialConcentration * 1000#
				Left_Title = Chr(181) & "g/L"
		End Select
		
		'Define Graph
		If (Number_Influent_Points = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.NumSets = 2
			numsets = 2
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.NumSets = 3
			numsets = 3
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.GraphType = 6 'Lines/Symbols
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.GraphStyle = 1 'Symbols

		'   grpBreak.ThisSet = 1
		'   grpBreak.NumPoints = Results.NPoints
		'   grpBreak.ThisSet = 2
		'   grpBreak.NumPoints = NData_Points

		' The following code where grpBreak.NumPoints is set is a rather
		' unfortunate kludge, in my opinion.  I could find no other way to
		' convince/force Visual Basic's graphical interface to accept two sets
		' of data that were of two different sizes, so I determined which one
		' was the smaller set and then filled the remainer of the smaller set
		' with copies of the last data point in it (X,Y) (note, the default
		' is for the data to hook back to the point (0,0) at the end of its
		' plotting due to the fact that, by default, the (X,Y) data points
		' that are unspecified are filled with 0's).
		' -- If possible, it would be nice to replace this with something
		' more elegant, but hey, it works. -- Eric J. Oman
		If (num_model_points > NData_Points) Then
			bigger = num_model_points
		Else
			bigger = NData_Points
		End If
		If (Number_Influent_Points > bigger) Then
			bigger = Number_Influent_Points
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisSet = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.NumPoints = bigger
		numpoints(0) = bigger
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisSet = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.NumPoints = bigger
		numpoints(1) = bigger
		If (Number_Influent_Points = 0) Then
			'Do nothing
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisSet = 3
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.NumPoints = bigger
			numpoints(2) = bigger
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.SymbolData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.SymbolData = 2 'triangle
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.SymbolData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.SymbolData = 6 'square

		If (Number_Influent_Points = 0) Then
			'Do nothing
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.SymbolData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.SymbolData = 8 'diamond
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ColorData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ColorData = 9 'Blue
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ColorData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ColorData = 12 'Red
		If (Number_Influent_Points = 0) Then
			'Do nothing
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ColorData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ColorData = 10 'Green
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.PatternData = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.PatternData = 1
		If (Number_Influent_Points = 0) Then
			'Do nothing
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.PatternData = 1
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.AutoInc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.AutoInc = 0 'No autoincrementation

		'**************************************************************
		'   grpBreak.ThisSet = 1
		'   For I = 1 To grpBreak.NumPoints
		'     grpBreak.ThisPoint = I
		'     If Cin(Component_Index, I) < 0 Then
		'       grpBreak.GraphData = 0#
		'     Else
		'       grpBreak.GraphData = Cin(Component_Index, I) 'Results.CP(Component, I)
		'     End If
		'     grpBreak.ThisPoint = I
		'     grpBreak.LabelText = ""
		'     grpBreak.ThisPoint = I
		'     grpBreak.XPosData = Td(I) * factor 'X_Values(I)
		'   Next I
		'   grpBreak.ThisPoint = 1
		'   grpBreak.LegendText = Trim$(Results.Component(Component_Index).Name)

		'---- I. Display Effluent Prediction
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim s As New Series 'first
		s.ChartType = SeriesChartType.Line

		'grpBreak.ThisSet = 1
		Select Case frmCompareData_WhichSet
			Case frmCompareData_WhichSet_PSDM
				For i = 1 To num_model_points
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.ThisPoint = i

					If (Results.CP(Component_Index, i) < 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.GraphData = 0#
						myY = 0

					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.GraphData = Results.CP(Component_Index, i) * c_factor
						myY = Results.CP(Component_Index, i) * c_factor
					End If
					''''grpBreak.LabelText = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.XPosData = Results.T(i) * t_factor 'X_Values(I)
					myX = Results.T(i) * t_factor 'X_Values(I)
					s.Points.AddXY(myX, myY)

				Next i
			Case frmCompareData_WhichSet_CPHSDM
				For i = 1 To num_model_points
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.ThisPoint = i
					If (CPM_Results.C_Over_C0(i) < 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.GraphData = 0#
						myY = 0
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.GraphData = CPM_Results.C_Over_C0(i) * c_factor
						myY = CPM_Results.C_Over_C0(i) * c_factor
					End If
					''''grpBreak.LabelText = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.XPosData = CPM_Results.T(i) * 24.0# * 60.0# * t_factor
					myX = CPM_Results.T(i) * 24.0# * 60.0# * t_factor
					s.Points.AddXY(myX, myY)

				Next i
		End Select

		s.LegendText = Trim$(Results.Component(Component_Index).Name) + " Prediction"
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisPoint = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.LegendText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.LegendText = "Effluent Prediction"
		'grpBreak.LegendText = Trim$(Results.Component(Component_Index).Name)

		'---- II. Display Effluent Data
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim e As New Series 'second
		e.ChartType = SeriesChartType.Point
		'grpBreak.ThisSet = 2
		For i = 1 To NData_Points
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisPoint = i
			If (C_Data_Points(Component_Index, i) < 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.GraphData = 0#
				myY = 0
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.GraphData = C_Data_Points(Component_Index, i) * c_factor
				myY = C_Data_Points(Component_Index, i) * c_factor
			End If
			''''grpBreak.LabelText = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.XPosData = T_Data_Points(i) * 24.0# * 60.0# * t_factor
			myX = T_Data_Points(i) * 24.0# * 60.0# * t_factor
			e.Points.AddXY(myX, myY)

		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisPoint = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.LegendText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.LegendText = "Effluent Data"

		e.LegendText = "Effluent Data"


		'---- III. Display Influent Data
		Dim f As New Series 'third
		f.ChartType = SeriesChartType.Line
		f.BorderDashStyle = ChartDashStyle.Dash
		'f.ChartType = SeriesChartType.Point
		If (Number_Influent_Points = 0) Then
			'Do nothing
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisSet = 3
			For i = 1 To Number_Influent_Points
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.ThisPoint = i
				If (C_Influent(Component_Index, i) < 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.GraphData = 0#
					myY = 0
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.GraphData = C_Influent(Component_Index, i) / comp_data.InitialConcentration * c_factor
					myY = C_Influent(Component_Index, i) / comp_data.InitialConcentration * c_factor
				End If
				''''grpBreak.LabelText = ""
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.XPosData = T_Influent(i) * t_factor
				myX = T_Influent(i) * t_factor
				f.Points.AddXY(myX, myY)
			Next i
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisPoint = 3
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.LegendText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.LegendText = "Influent Data"
		End If

		f.LegendText = "Influent Data"



		'---- Run the kludge mentioned above.
		If (bigger > NData_Points) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisSet = 2
			LastPointI = NData_Points

			'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SameX = T_Data_Points(LastPointI) * 24# * 60# * t_factor
			SameY = C_Data_Points(Component_Index, LastPointI) * c_factor
			For i = LastPointI + 1 To bigger
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.ThisPoint = i
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.GraphData = SameY
				''''grpBreak.ThisPoint = i
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.XPosData = SameX

				e.Points.AddXY(SameX, SameY)

			Next i
		End If
		Select Case frmCompareData_WhichSet
			Case frmCompareData_WhichSet_PSDM
				If (bigger > num_model_points) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.ThisSet = 1
					LastPointI = num_model_points
					'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SameX = Results.T(LastPointI) * t_factor
					SameY = Results.CP(Component_Index, LastPointI) * c_factor
					For i = LastPointI + 1 To bigger
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.GraphData = SameY
						''''grpBreak.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.XPosData = SameX
						s.Points.AddXY(SameX, SameY)
					Next i
				End If
			Case frmCompareData_WhichSet_CPHSDM
				If (bigger > num_model_points) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.ThisSet = 1
					LastPointI = num_model_points
					'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SameX = CPM_Results.T(LastPointI) * 24# * 60# * t_factor
					SameY = CPM_Results.C_Over_C0(LastPointI) * c_factor
					For i = LastPointI + 1 To bigger
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.GraphData = SameY
						''''grpBreak.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.XPosData = SameX
						s.Points.AddXY(SameX, SameY)
					Next i
				End If
		End Select
		If (Number_Influent_Points = 0) Then
			'Do nothing
		Else
			If (bigger > Number_Influent_Points) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.ThisSet = 3
				LastPointI = Number_Influent_Points
				'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SameX = T_Influent(LastPointI) * t_factor
				SameY = C_Influent(Component_Index, LastPointI) / Component(Component_Index).InitialConcentration * c_factor
				For i = LastPointI + 1 To bigger
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.ThisPoint = i
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.GraphData = SameY
					''''grpBreak.ThisPoint = i
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object SameX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.XPosData = SameX
					f.Points.AddXY(SameX, SameY)
				Next i
			End If
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternedLines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.PatternedLines = 0
		Data_Max = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'


		Chart1.Series.Add(s)
		Chart1.Series.Add(e)
		Chart1.Series.Add(f)


		Chart1.ChartAreas(0).AxisY.Minimum = 0
		Chart1.ChartAreas(0).AxisX.Minimum = 0
		'Chart1.ChartAreas(0).AxisY.Maximum = (Int(Data_Max * 10.0# + 1)) / 10.0#
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisTicks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisTicks = 4
		'grpBreak.GridStyle = 0

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisStyle = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisMin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisMin = 0#
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.BottomTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.BottomTitle = Bottom_Title
		Chart1.ChartAreas(0).AxisX.Title = Bottom_Title

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.LeftTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.LeftTitle = Left_Title
		Chart1.ChartAreas(0).AxisY.Title = Left_Title
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.DrawMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.DrawMode = 2

	End Sub
	
	
	'UPGRADE_WARNING: Form event frmModelDataComparison.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmModelDataComparison_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If UnloadMe Then Me.Close()
	End Sub
	Private Sub frmModelDataComparison_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		Dim J, i As Short
		Me.Text = frmCompareData_caption
		Call Populate_Scrollboxes()
		Call CenterOnForm(Me, frmMain)
		''''Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmShow_Data_And_Prediction.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmShow_Data_And_Prediction.Height / 2)
		'If Obj_Function() Then
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		Call Draw_Curves(cboCompo.SelectedIndex + 1)
		Call cboGraphType_SelectedIndexChanged(cboGraphType, New System.EventArgs()) 'Wake cboGraphType up
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		UnloadMe = False
		Call cboGrid_SelectedIndexChanged(cboGrid, New System.EventArgs())
		'Else
		'  UnloadMe = True
		'End If
	End Sub
	'UPGRADE_WARNING: Event frmModelDataComparison.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmModelDataComparison_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize

		rs.ResizeAllControls(Me)

		'If WindowState = 1 Then
		'  frmPFPSDM.WindowState = 1
		'  frmPlantData.WindowState = 1
		'End If
	End Sub
	Private Sub frmModelDataComparison_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call UserPrefs_Save()
	End Sub
	
	
	Private Sub UserPrefs_Load()
		Dim X As Integer
		On Error GoTo err_DATAPRED_UserPrefs_Load
		X = CInt(INI_Getsetting("DATAPRED_cboCUnits"))
		If ((X >= 0) And (X <= cboCUnits.Items.Count - 1)) Then
			cboCUnits.SelectedIndex = X
		End If
		X = CInt(INI_Getsetting("DATAPRED_cboTUnits"))
		If ((X >= 0) And (X <= cboTUnits.Items.Count - 1)) Then
			cboTUnits.SelectedIndex = X
		End If
		X = CInt(INI_Getsetting("DATAPRED_cboGraphType"))
		If ((X >= 0) And (X <= cboGraphType.Items.Count - 1)) Then
			cboGraphType.SelectedIndex = X
		End If
		X = CInt(INI_Getsetting("DATAPRED_cboGrid"))
		If ((X >= 0) And (X <= cboGrid.Items.Count - 1)) Then
			cboGrid.SelectedIndex = X
		End If
		Exit Sub
resume_err_DATAPRED_UserPrefs_Load: 
		Call UserPrefs_Save()
		Exit Sub
err_DATAPRED_UserPrefs_Load: 
		Resume resume_err_DATAPRED_UserPrefs_Load
	End Sub
	Private Sub UserPrefs_Save()
		Dim X As Integer
		X = cboCUnits.SelectedIndex
		Call INI_PutSetting("DATAPRED_cboCUnits", Trim(CStr(X)))
		X = cboTUnits.SelectedIndex
		Call INI_PutSetting("DATAPRED_cboTUnits", Trim(CStr(X)))
		X = cboGraphType.SelectedIndex
		Call INI_PutSetting("DATAPRED_cboGraphType", Trim(CStr(X)))
		X = cboGrid.SelectedIndex
		Call INI_PutSetting("DATAPRED_cboGrid", Trim(CStr(X)))
	End Sub

	Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
		'Me.Close()    'not working after reopening
		Me.Dispose()   'Shang 
	End Sub

	Private Sub cmdPrint_Click(sender As Object, e As EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()

		'	'''Dim i As Integer
		'		'''Dim H As Single
		'		'''Dim W As Single
		'		'''  'printer.ScaleLeft = -1080  'Set a 3/4-inch margin
		'		'''  'printer.ScaleTop = -1080
		'		'''  'printer.CurrentX = 0
		'		'''  'printer.CurrentY = 0
		'		'''  '
		'		'''  '  printer.FontSize = 12
		'		'''  '  printer.FontBold = True
		'		'''  '  printer.FontUnderline = True
		'		'''  '  printer.Print "Input data for the Plug-Flow Pore And Surface Diffusion Model"
		'		'''  '  printer.FontSize = 10
		'		'''  '  printer.FontBold = False
		'		'''  '  printer.FontUnderline = False
		'		'''  '  '-- Print Filename
		'		'''  '  printer.Print
		'		'''  '  printer.Print "From Data File: "; Filename
		'		'''  '---- Print the graph ------------------------
		'		'''  For i = 1 To Number_Component
		''		'''    grpBreak.ThisPoint = i
		'	'''    grpBreak.PatternData = i - 1
		'	'''  Next i
		'	'''  H = grpBreak.Height
		'	'''  W = grpBreak.Width
		'	'''  grpBreak.Visible = False 'Hide it before printing
		'	'''  If (Printer.Width < Printer.Height) Then
		'	'''    grpBreak.Height = CSng(Printer.Height / 2#)
		'	'''    grpBreak.Width = Printer.Width
		'	'''  Else
		'	'''    grpBreak.Height = Printer.Height
		'	'''    grpBreak.Width = Printer.Width
		'	'''  End If
		'	'''  grpBreak.PrintStyle = 2
		'	'''  grpBreak.DrawMode = 5
		'	'''  grpBreak.Height = H
		'	'''  grpBreak.Width = W
		'	'''  grpBreak.Visible = True
		'	'''  grpBreak.PrintStyle = 2
		'	'''  grpBreak.DrawMode = 2
		'	'''  Printer.EndDoc

	End Sub





	'Private Function Obj_Function() As Integer
	'Dim i, J As Integer
	'Dim ncomp As Long, ndata  As Long, np As Long, temp As String, Error_Code As Integer
	'ReDim Fmin(Results.NComponent) As Double
	'ReDim TP(Results.npoints) As Double, CP(Results.NComponent, Results.npoints) As Double
	'ReDim Td(NData_Points) As Double, Cd(Results.NComponent, NData_Points) As Double, Cin(Results.NComponent, NData_Points) As Double
	'
	'  ncomp = CLng(Results.NComponent)
	'  ndata = CLng(NData_Points)
	'  np = CLng(Results.npoints)
	'
	'  For i = 1 To Results.npoints
	'    TP(i) = Results.T(i)
	'    For J = 1 To Results.NComponent
	'      CP(J, i) = Results.CP(J, i)
	'    Next J
	'  Next i
	'  For i = 1 To NData_Points
	'    Td(i) = T_Data_Points(i) * 24# * 60#
	'    For J = 1 To Results.NComponent
	'      Cd(J, i) = C_Data_Points(J, i)
	'    Next J
	'  Next i
	'
	'On Error GoTo Error_In_OBJFUN
	'  'Call OBJFUN(ncomp, ndata, np, Tp(1), CP(1, 1), Td(1), Cd(1, 1), Cin(1, 1), Fmin(1))
	'  Obj_Function = True
	'  Exit Function
	'
	'Error_In_OBJFUN:
	'  Error_Code = Err
	'  temp = "Error " & Format$(Error_Code, "0") & " : " & Error$(Error_Code)
	'  MsgBox "Fatal Error with OBJFUN.DLL. Calculations Stoppped." & Chr$(13) & temp, MB_ICONEXCLAMATION, AppName_For_Display_long
	'  Obj_Function = False
	'  Resume Exit_Obj_Function
	'Exit_Obj_Function:
	'End Function
End Class