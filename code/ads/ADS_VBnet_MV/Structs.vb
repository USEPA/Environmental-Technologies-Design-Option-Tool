Option Strict Off
Option Explicit On
Module Structs
	
	
	Public Const USE_GASPHASE_WAKAO_AND_FUNAZUKRI As Boolean = False 'False 'True
	
	''''Global Const Application_Name = "AdXsorption Design Software"
	''''Global Const name_app = "AdXDesignS"
	''''Global Const name_app_long = "AdXsorption Design Software"
	
	'Constants
	'Global Const NVersion = 1.3
	Public Const NVersion As Double = 1.4
	Public Const Latest_DataVersion_Major As Short = 1
	Public Const Latest_DataVersion_Minor As Short = 60
	Public Const Number_Compo_Max As Short = 10
	Public Const MAXCHEMICAL As Short = Number_Compo_Max
	Public Const Number_Compo_Max_PFPSDM As Short = 6
	Public Const Number_Compo_Max_ECM As Short = 9
	Public Const Number_Compo_Max_CPM As Short = 1
	Public Const Number_Points_Max As Short = 400
	Public Const Number_Data_Points_Max As Short = 400
	Public Const Number_Max_Influent_Points As Short = 400
	Public Const Max_Number_Correlation_Compo As Short = 25
	Public Const Max_Number_Water_Correlations As Short = 25
	Public Const CPM_Max_Points As Short = 100
	Public Const PI As Double = 3.14159265359
	
	'Modified Hokanson 2/8/97
	'Global Const Max_Radial_Collocation = 6
	Public Const Max_Radial_Collocation As Short = 18
	Public Const Max_Equations_DGEAR As Short = 750
	'end Modified Hokanson 2/8/97
	
	Public Const Max_Axial_Collocation As Short = 18
	Public Const Max_Number_Fouling_Iterations As Short = 100
	Public Const Maximum_Beds_In_Series As Short = 200
	Public Const EPS_ERROR_CRITERIA As Double = 0.0005
	
	Structure Tempo_Data
		Dim MW As Double
		Dim Solubility As Double
		Dim Density As Double
		Dim R_Index As Double
		Dim T As Double
		Dim Pvap As Double
	End Structure
	
	Structure IPES_Input
		Dim Adsorbent As String
		Dim BB As Double
		Dim W0 As Double
		Dim RH As Double 'Relative Humidity
		Dim C As Double 'Concentration
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(6),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=6)> Public Phase() As Char
		Dim IMOD As Short
		Dim GM As Double
		Dim OMAG As Double
		Dim NL As Short
	End Structure
	
	Structure IPES_Output
		Dim XN As Double
		Dim XK1 As Double
		Dim XK2 As Double
		Dim CSAV As Double
		Dim QSAV As Double
		Dim CBEG As Double
		Dim CEND As Double
		Dim RSQD As Double
		Dim RMSE As Double
		<VBFixedArray(30)> Dim Error_Matrix() As Short
		<VBFixedArray(200)> Dim CorrelationPoints_lnC() As Double
		<VBFixedArray(200)> Dim CorrelationPoints_lnQ() As Double
		Dim CorrelationPoints_NumPoints As Short
		<VBFixedArray(200)> Dim QCAP() As Double
		<VBFixedArray(200)> Dim ADSP() As Double
		<VBFixedArray(200)> Dim PI() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If Error_Matrix Is Nothing Then ReDim Error_Matrix(30)
			'UPGRADE_WARNING: Lower bound of array CorrelationPoints_lnC was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If CorrelationPoints_lnC Is Nothing Then ReDim CorrelationPoints_lnC(200)
			'UPGRADE_WARNING: Lower bound of array CorrelationPoints_lnQ was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If CorrelationPoints_lnQ Is Nothing Then ReDim CorrelationPoints_lnQ(200)
			'UPGRADE_WARNING: Lower bound of array QCAP was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If QCAP Is Nothing Then ReDim QCAP(200)
			'UPGRADE_WARNING: Lower bound of array ADSP was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If ADSP Is Nothing Then ReDim ADSP(200)
			'UPGRADE_WARNING: Lower bound of array PI was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If PI Is Nothing Then ReDim PI(200)
		End Sub
	End Structure
	
	Structure IPES_Variable
		'UPGRADE_NOTE: Input was upgraded to Input_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Input_Renamed As IPES_Input
		'UPGRADE_WARNING: Arrays in structure Output may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Output As IPES_Output
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			Output.Initialize()
		End Sub
	End Structure
	
	Structure Correlation_Compound_Type
		Dim Name As String
		<VBFixedArray(2)> Dim Coeff() As Double
		'Dim Coeff(2) As Double


		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If Coeff Is Nothing Then ReDim Coeff(2)
		End Sub
	End Structure
	
	Structure Correlation_Water_Type
		Dim Name As String
		Dim Coeff() As Double

		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If Coeff Is Nothing Then ReDim Coeff(4)
		End Sub
	End Structure
	
	Public Const KNSOURCE_ISOTHERMDB As Short = 1
	Public Const KNSOURCE_IPES As Short = 2
	Public Const KNSOURCE_USERINPUT As Short = 3
	
	
	Public Const IPESMETHOD_LIQ_3PARAM As Short = 1
	Public Const IPESMETHOD_LIQ_DRUNIFORM As Short = 2
	Public Const IPESMETHOD_LIQ_DRNONUNIFORM As Short = 3
	Public Const IPESMETHOD_GAS_DRZERORH As Short = 101
	Public Const IPESMETHOD_GAS_CALGONBPL As Short = 102
	Public Const IPESMETHOD_GAS_DRSPREADINGP As Short = 103
	
	
	Structure ComponentPropertyType
		'Isotherm Freundlich q=k*C^OneOverN
		'***** Properties: *****
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		'	<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Name() As Char
		Public Name As String   'Shang 
		Dim CAS As Integer
		Dim MW As Double 'g/mol
		Dim MolarVolume As Double 'cm3/mol
		Dim BP As Double 'degrees Celcius
		Dim InitialConcentration As Double 'mg/l
		Dim Use_K As Double '(mg/g)*(L/mg)^OneoverN
		Dim Use_OneOverN As Double '(-)
		Dim Source_KandOneOverN As Short 'Isotherm DB / IPES / User-Input
		Dim UserEntered_K As Double
		Dim UserEntered_OneOverN As Double
		Dim Treatment_Objective As Double
		'***** IPES RELATED: *****
		Dim IPES_OrderOfMagnitude As Double
		Dim IPES_NumRegressionPts As Short
		Dim IPES_RelativeHumidity As Double
		Dim IPES_EstimationMethod As Short
		Dim Liquid_Density As Double 'g/cm^3
		Dim Aqueous_Solubility As Double 'mg/L
		Dim Vapor_Pressure As Double 'Pa
		Dim Refractive_Index As Double '(-)
		Dim IPESResult_K As Double
		Dim IPESResult_OneOverN As Double
		'***** Isotherm Database: *****
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(70),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=70)> Public IsothermDB_Component_Name() As Char
		'(note: includes CAS and Name exactly as it appears on the Freundlich Isotherm Parameters form)
		Dim IsothermDB_Range_Num As Short
		'(this is the list index [1-n] of the selected range.)
		Dim IsothermDB_K As Double
		Dim IsothermDB_OneOverN As Double
		'***** Kinetics *****
		Dim SPDFR As Double
		Dim SPDFR_Low_Concentration As Short
		Dim Use_SPDFR_Correlation As Short
		<VBFixedArray(3)> Dim Corr() As Short
		Dim kf As Double 'cm/s
		Dim Ds As Double 'cm2/s
		Dim Dp As Double 'cm2/s
		<VBFixedArray(3)> Dim KP_User_Input() As Double
		Dim Tortuosity As Double
		Dim Use_Tortuosity_Correlation As Short
		Dim Constant_Tortuosity As Short
		'***** ECM *****
		Dim K_Reduction As Short 'Boolean
		'UPGRADE_WARNING: Arrays in structure Correlation may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Correlation As Correlation_Compound_Type
		'Correlation to calculate K reduction
		Dim Is_Selected_On_List As Boolean
		'TEMPORARY INTERNAL VARIABLE: NOT SAVED.
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If Corr Is Nothing Then ReDim Corr(3)
			If KP_User_Input Is Nothing Then ReDim KP_User_Input(3)
			Correlation.Initialize()
		End Sub
	End Structure
	
	Structure PropertyUnitsType
		Dim MW As String
		Dim MolarVolume As String
		Dim BP As String
		Dim InitialConcentration As String
		Dim Liquid_Density As String
		Dim Aqueous_Solubility As String
		Dim Vapor_Pressure As String
		Dim Refractive_Index As String
		Dim k As String
		Dim BedTemperature As String
		Dim BedPressure As String
		Dim BedFluidDensity As String
		Dim BedFluidViscosity As String
	End Structure
	Public PropertyUnits As PropertyUnitsType
	
	Structure BedPropertyType
		Dim length As Double 'm
		Dim Diameter As Double 'm
		Dim Weight As Double 'kg
		Dim Flowrate As Double 'm3/s
		Dim Density As Double 'g/cm3
		Dim SuperficialVelocity As Double 'm/s
		Dim Porosity As Double '(-)
		Dim InterstitialVelocity As Double 'm/s
		Dim Area As Double 'm2
		Dim Volume As Double 'm3
		Dim TAU As Double 'min (packed bed contact time)
		Dim NumberOfBeds As Short '(-)
		Dim WaterDensity As Double 'g/cm3
		Dim WaterViscosity As Double 'g/cm.s
		Dim Temperature As Double 'C
		Dim Pressure As Double 'Atm
		Dim Phase As Short '=0 -> liquid, =1 -> gas
		'UPGRADE_WARNING: Arrays in structure Water_Correlation may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Water_Correlation As Correlation_Water_Type 'Correlation to calcualte K reduc
		Dim UnitsLength As Short
		Dim UnitsDiameter As Short
		Dim UnitsWeight As Short
		Dim UnitsFlowrate As Short
		Dim UnitsEBCT As Short
		Dim UnitsFluidDensity As Short
		Dim UnitsFluidViscosity As Short
		Dim UnitsFluidTemperature As Short
		Dim UnitsFluidPressure As Short
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			Water_Correlation.Initialize()
		End Sub
	End Structure
	
	Structure CarbonPropertyType
		'Manu As String
		Dim Name As String
		Dim Porosity As Double ' -
		Dim Density As Double 'g/cm3  Apparent density
		Dim ParticleRadius As Double 'm
		Dim Tortuosity As Double ' -            'UNUSED!!!
		' Need to add W0, BB, Polanyi Exponent here!
		' It appears safe to add these variables.
		Dim W0 As Double
		Dim BB As Double
		Dim PolanyiExponent As Double
		'---- Added by EJO on 11/1/96 for kf (external mass xfer coefficient)
		Dim ShapeFactor As Double ' -
	End Structure
	
	Structure DataCarbon
		Dim Index As Short
		Dim CAS As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Name() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public NameC() As Char
		Dim pHmin As Double
		Dim pHmax As Double
		Dim k As Double
		Dim N As Double 'Actually, this is 1/n
		Dim Cmin As Double
		Dim Cmax As Double
		Dim Temperature As Double
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Phase() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Source() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Comments() As Char
	End Structure
	
	Structure Para_Int
		Dim Init As Double
		'UPGRADE_NOTE: End was upgraded to End_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim End_Renamed As Double
		Dim np As Short
		'UPGRADE_NOTE: Step was upgraded to Step_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Step_Renamed As Double
	End Structure
	
	Structure Throughput
		Dim T As Double
		Dim C As Double
	End Structure
	
	Structure ResultsType
		Dim NComponent As Short
		Dim npoints As Short
		<VBFixedArray(Number_Compo_Max, Number_Points_Max)> Dim CP(, ) As Double
		<VBFixedArray(Number_Points_Max)> Dim T() As Double
		<VBFixedArray(Number_Compo_Max)> Dim ThroughPut_05() As Throughput
		<VBFixedArray(Number_Compo_Max)> Dim ThroughPut_50() As Throughput
		<VBFixedArray(Number_Compo_Max)> Dim ThroughPut_95() As Throughput
		'UPGRADE_WARNING: Array Component may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
		<VBFixedArray(Number_Compo_Max)> Dim Component() As ComponentPropertyType
		Dim Bed As BedPropertyType
		Dim Carbon As CarbonPropertyType
		Dim Use_Tortuosity_Correlation As Short
		Dim Constant_Tortuosity As Short
		<VBFixedArray(Number_Compo_Max)> Dim NumPoints_Before_ThroughPut_100() As Short 'Used by PSDM to cut off display when C/C0 >= 1
		Dim is_psdm_in_room_model As Short
		Dim int_Which_PSDMR_Model As Short
		<VBFixedArray(Number_Compo_Max)> Dim psdmroom_Crss() As Double 'ug/L
		Dim AnyCrCloseToZero As Short
		Shared Initialized As Boolean = False   'Shang to add a flag
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If CP Is Nothing Then ReDim CP(Number_Compo_Max, Number_Points_Max)
			If T Is Nothing Then ReDim T(Number_Points_Max)
			If ThroughPut_05 Is Nothing Then ReDim ThroughPut_05(Number_Compo_Max)
			If ThroughPut_50 Is Nothing Then ReDim ThroughPut_50(Number_Compo_Max)
			If ThroughPut_95 Is Nothing Then ReDim ThroughPut_95(Number_Compo_Max)
			If Component Is Nothing Then ReDim Component(Number_Compo_Max)
			Bed.Initialize()
			If NumPoints_Before_ThroughPut_100 Is Nothing Then ReDim NumPoints_Before_ThroughPut_100(Number_Compo_Max)
			'UPGRADE_WARNING: Lower bound of array psdmroom_Crss was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If psdmroom_Crss Is Nothing Then ReDim psdmroom_Crss(Number_Compo_Max)
		End Sub
	End Structure
	
	Structure PSDMInputsType
		<VBFixedArray(15)> Dim VARS1() As Double
		<VBFixedArray(Number_Compo_Max, 19)> Dim VARS2(, ) As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array VARS1 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If VARS1 Is Nothing Then ReDim VARS1(15)
			'UPGRADE_WARNING: Lower bound of array VARS2 was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If VARS2 Is Nothing Then ReDim VARS2(Number_Compo_Max, 19)
        End Sub
	End Structure
	
	Structure ECM_Data
		Dim Index As Short
		Dim Bed_Volume_Fed As Double
		Dim Wave_Velocity As Double
		Dim Dimensionless_Bed_Length As Double
		Dim SS_Treatment_Capacity As Double
		<VBFixedArray(Number_Compo_Max)> Dim Solid_Concentration() As Double
		<VBFixedArray(Number_Compo_Max)> Dim Liquid_Concentration() As Double
		<VBFixedArray(Number_Compo_Max)> Dim C_Over_C0() As Double
		Dim Carbon_Usage_Rate As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If Solid_Concentration Is Nothing Then ReDim Solid_Concentration(Number_Compo_Max)
			If Liquid_Concentration Is Nothing Then ReDim Liquid_Concentration(Number_Compo_Max)
			If C_Over_C0 Is Nothing Then ReDim C_Over_C0(Number_Compo_Max)
		End Sub
	End Structure
	
	Structure ECM_MASSBAL
		<VBFixedArray(Number_Compo_Max)> Dim MASSBAL_C0_e_Vf() As Double
		<VBFixedArray(Number_Compo_Max)> Dim MASSBAL_TERM_SUM() As Double
		<VBFixedArray(Number_Compo_Max)> Dim MASSBAL_PERCENT_ERR() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If MASSBAL_C0_e_Vf Is Nothing Then ReDim MASSBAL_C0_e_Vf(Number_Compo_Max)
			If MASSBAL_TERM_SUM Is Nothing Then ReDim MASSBAL_TERM_SUM(Number_Compo_Max)
			If MASSBAL_PERCENT_ERR Is Nothing Then ReDim MASSBAL_PERCENT_ERR(Number_Compo_Max)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure Output_ECM_MASSBAL may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public Output_ECM_MASSBAL As ECM_MASSBAL
	
	Structure CPM_Data
		<VBFixedArray(CPM_Max_Points)> Dim T() As Double
		<VBFixedArray(CPM_Max_Points)> Dim C_Over_C0() As Double
		<VBFixedArray(7)> Dim Par() As Double
		Dim ThroughPut_05 As Throughput
		Dim ThroughPut_50 As Throughput
		Dim ThroughPut_95 As Throughput
		'UPGRADE_WARNING: Arrays in structure Component may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Component As ComponentPropertyType
		Dim Bed As BedPropertyType
		Dim Carbon As CarbonPropertyType
		Dim Use_Tortuosity_Correlation As Short
		Dim Constant_Tortuosity As Short
		Shared Initialized As Boolean = False   'Shang added

		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			If T Is Nothing Then ReDim T(CPM_Max_Points)
			If C_Over_C0 Is Nothing Then ReDim C_Over_C0(CPM_Max_Points)
			If Par Is Nothing Then ReDim Par(7)
			Component.Initialize()
			Bed.Initialize()
		End Sub
	End Structure
	
	Structure Isotherm_Data
		Dim number As Short
		Dim Selected As Integer
		Dim Record As DataCarbon
	End Structure
	
	Structure Isotherm_Data_Save
		Dim New_CAS As Short
		Dim New_Name As Short
		Dim Record As DataCarbon
	End Structure
	
	Structure Isotherm_Chemical
		Dim CAS As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Name() As Char
		Dim Update_Name As Short
		Dim Update_CAS As Short
	End Structure
	
	Structure Carbon_Data
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Name() As Char
		Dim Density As Double
		Dim ParticleRadius As Double
		Dim Porosity As Double
		Dim ShapeFactor As Double
		Dim Tortuosity As Double 'UNUSED!
		Dim Phase As Short '1=Liquid, 2 =gas
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Type() As Char
		Dim W0 As Double
		Dim b As Double
		Dim PolanyiExponent As Double
	End Structure
	
	'Variables for Help System
	Public HelpFile As String
	
	'Variable for security
	Public Open_Database As Short
	
	'Variables
	Public PFPSDM_Path As String
	Public Error_In_Kinetic_Calculation As Short
	Public Database_Path As String
	Public Exe_Path As String 'NEW 9/2/98.
	Public Flag_Openfile As Short
	Public Update_Value_From_Carbon, Update_Value_From_IPES As Short
	Public Temp_Text As String 'Temporary string to store former value of a text box
	Public AddFlag As Short 'Flag - True = Add a chemical to the list - False = Edit chemical properties
	Public Component_Number_Selected As Short
	Public State_Check_Water(2) As Short 'Flag to know whether correlations are used for water properties
	Public Filename, Previous_FileName As String
	'UPGRADE_WARNING: Arrays in structure Results may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public Results As ResultsType
	'UPGRADE_WARNING: Arrays in structure PSDM_Inputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public PSDM_Inputs As PSDMInputsType
	Public Print_To_Printer As Short 'Flag - True = print to Printer - False = Print to file
	Public NL As String
	Public batchrun As Short
	Public Treatment_Objective(Number_Compo_Max_PFPSDM) As Throughput
	
	'Variable for Tortuosity as a function of Time
	Public Use_Tortuosity_Correlation, Constant_Tortuosity As Short
	
	'Variable for the search function
	Public Start_Search As Short
	Public Find_String As String
	Public Index_NameT, Index_Find As Short
	
	'Variables for K reduction
	Public Number_Correlations_Compounds As Short
	Public Number_Water_Correlations As Short
	'UPGRADE_WARNING: Array Correlations_For_Classes may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
	Public Correlations_For_Classes(Max_Number_Correlation_Compo) As Correlation_Compound_Type
	'UPGRADE_WARNING: Array Correlations_For_Water may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
	Public Correlations_For_Water(Max_Number_Water_Correlations) As Correlation_Water_Type
	
	'Variable used to transfer data to excel
	Public Excel_4 As Short
	Public PFPSDM_Excel As Short
	
	'Variables used to edit the isotherm database
	Public Mode_Chemical, Mode_Isotherm As Short
	Public Iso_Data As Isotherm_Data
	Public Iso_Data_Save As Isotherm_Data_Save
	Public Iso_Chemical As Isotherm_Chemical
	
	'Variables used to edit the carbon database
	Public Mode_Manu As Short
	Public Mode_Carbon As Short
	Public Name_Manufacturer_In As String
	Public Name_Manufacturer_Out As String
	Public Carbon_Data_In As Carbon_Data
	Public Carbon_Data_Out As Carbon_Data
	
	'Variables for ECM model
	'UPGRADE_WARNING: Array Output_ECM may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
	Public Output_ECM(Number_Compo_Max) As ECM_Data
	Public Number_Component_ECM As Short
	Public Component_Index_ECM(Number_Compo_Max_ECM) As Object
	
	'Variables for the Contant Pattern Model
	Public CPHSDM_Excel As Short
	Public Number_Component_CPM As Short
	Public Component_Index_CPM As Short
	'UPGRADE_WARNING: Arrays in structure CPM_Results may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public CPM_Results As CPM_Data
	
	'Variables for the PFPSDM FORTRAN program
	Public C_Influent(Number_Compo_Max, Number_Max_Influent_Points) As Double
	Public T_Influent(Number_Max_Influent_Points) As Double
	Public Number_Influent_Points As Short
	Public TimeP As Para_Int
	Public MC, NC As Short
	Public Number_Component As Short
	Public Bed As BedPropertyType
	'UPGRADE_WARNING: Array Component may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
	Public Component(Number_Compo_Max) As ComponentPropertyType 'Component(0) is a component used for temporary storage. Component(1) to Component(N) are the components in the listbox
	Public Carbon As CarbonPropertyType
	
	Public Component_Index_PFPSDM(Number_Compo_Max_PFPSDM) As Short
	Public Number_Component_PFPSDM As Short
	
	Public IPES_Data As IPES_Variable
	Public Properties_For_IPES As Tempo_Data
	
	'Flags to avoid problems when loading a file...
	Public Use_Update_Display_Kinetic As Short
	
	'Flags to check whether or not the user wants the data to be saved before exiting
	' in the QueryUnload event from the main window
	Public ReallyQuit As Short
	
	'Arrays to store the experimental data points for comparison to a PSDM simulation
	Public T_Data_Points(Number_Data_Points_Max) As Double
	Public C_Data_Points(Number_Compo_Max, Number_Data_Points_Max) As Double
	Public NData_Points As Short
	'Global NComponents As Integer
	
	'Variable to store isotherm parameters to plot it
	Public IsothermProperties As DataCarbon
	
	'Variable to tell frmBatch its default model to simulate
	Public BatchSimulation_DefaultModel As Short
	
	'Variables to tell frmStEPPImport what to do.
	Public Const STEPPIMPORT_ADDCOMPONENTS As Short = 1
	Public Const STEPPIMPORT_IPESCOMPONENT As Short = 2
	Public StEPPImportDestination As Short
	Public StEPPImportSuccess As Short
	Structure StEPP_to_IPES_Properties_type
		Dim Name As String
		Dim MW As Double
		Dim MolarVolume As Double
		Dim BP As Double
	End Structure
	Public StEPP_to_IPES_Properties As StEPP_to_IPES_Properties_type
End Module