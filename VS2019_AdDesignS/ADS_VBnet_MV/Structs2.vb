Option Strict Off
Option Explicit On
Module Structs2
	
	'//////// COMMUNICATIONS WITH frmEditAdsorber: ////////////////////////////////////////////////////////
	Structure rec_adsorber_db_manufacturers
		Dim UniqueID As String 'Must be an integer in string form!
		Dim Name As String
	End Structure
	
	Structure rec_adsorber_db_adsorbers
		Dim UniqueID_Manufacturer As Short
		Dim Phase As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public PartNumber() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public InternalArea() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public MaxCapacity() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public OutsideDiameter() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public DesignPressure() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public DesignFlowRange() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public DefaultFlowRate() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(100),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=100)> Public Note() As Char
	End Structure
	
	Public adsorber_db_num_manufacturers As Short
	Public adsorber_db_manufacturers() As rec_adsorber_db_manufacturers
	
	Public adsorber_db_num_adsorbers As Short
	Public adsorber_db_adsorbers() As rec_adsorber_db_adsorbers
	
	Structure rec_frmEditAdsorber_ReturnParameters
		Dim D As Double
		Dim L As Double
		Dim M As Double
		Dim Q As Double
	End Structure
	
	Public frmEditAdsorber_ReturnParameters As rec_frmEditAdsorber_ReturnParameters
	
	
	'//////// COMMUNICATIONS WITH frmEditAdsorberData: /////////////////////////////////////////////////
	Public frmEditAdsorberData_Record As rec_adsorber_db_adsorbers
	
	
	'//////// COMMUNICATIONS WITH frmEditCarbonData: /////////////////////////////////////////////////
	Structure frmEditCarbonData_Record_Type
		Dim PhaseIsLiquid As Boolean
		Dim Name As String
		Dim Manufacturer As String
		Dim AppDen As Double
		Dim ParticleRadius As Double
		Dim ParticlePorosity As Double
		Dim AdsType As String
		Dim W0 As Double
		Dim BB As Double
		Dim PolanyiExponent As Double
	End Structure
	Public frmEditCarbonData_Record As frmEditCarbonData_Record_Type
	
	
	'//////// COMMUNICATIONS WITH frmEditIsothermData: /////////////////////////////////////////////////
	Structure frmEditIsothermData_Record_Type
		Dim PhaseIsLiquid As Boolean
		Dim Name As String
		Dim k As Double
		Dim OneOverN As Double
		Dim Cmin As Double
		Dim Cmax As Double
		Dim pHmin As Double
		Dim pHmax As Double
		Dim Source As String
		Dim CarbonName As String
		Dim Tmin As String
		Dim CAS As String
		Dim Comments As String
	End Structure
	Public frmEditIsothermData_Record As frmEditIsothermData_Record_Type
	
	
	
	'---- frmConcentrations variables
	Public frmConcentrations_cancelled As Short
	'UPGRADE_WARNING: Lower bound of array frmConcentrations_Times was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public frmConcentrations_Times(400) As Double
	'UPGRADE_WARNING: Lower bound of array frmConcentrations_Concs was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public frmConcentrations_Concs(10, 400) As Double
	Public frmConcentrations_NumPoints As Short
	Public frmConcentrations_NumConcs As Short
	Public frmConcentrations_caption As String
	Public frmConcentrations_TimeOrderImportant As Short
	Public frmConcentrations_Cunits As String
	Public frmConcentrations_Tunits As String
	
	'---- frmShow_Data_And_Prediction variables
	Public frmCompareData_WhichSet As Short
	Public Const frmCompareData_WhichSet_PSDM As Short = 1
	Public Const frmCompareData_WhichSet_CPHSDM As Short = 2
	Public frmCompareData_caption As String
	
	
	
	
	
	
	'MISCELLANEOUS.
	Public FileNote As String
End Module