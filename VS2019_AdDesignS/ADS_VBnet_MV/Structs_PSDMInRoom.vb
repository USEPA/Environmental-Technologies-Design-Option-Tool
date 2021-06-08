Option Strict Off
Option Explicit On
Module Structs_PSDMInRoom
	
	'
	' NOTE:
	' IF Distribute_PSDMInRoom IS SET TO FALSE, THEN
	' Activate_PSDMInRoom IS SET TO FALSE.
	' IF Distribute_PSDMInRoom IS SET TO TRUE, THEN
	' THE EXISTENCE OF THE FILE $(AppPath)\PSDMROOM.DAT IS
	' CHECKED.  IF THIS FILE EXISTS, Activate_PSDMInRoom
	' IS SET TO TRUE.  IF THIS FILE DOES NOT EXIST,
	' Activate_PSDMInRoom IS SET TO FALSE.
	'
	' IF Activate_PSDMInRoom IS SET TO TRUE, THE PSDMR
	' MENU ENTRIES AND FILE-LOAD CAPABILITIES ARE ENABLED;
	' IF SET TO FALSE, THESE CAPABILITIES ARE DISABLED.
	'
	Public Const Distribute_PSDMInRoom As Boolean = True
	'Global Const Distribute_PSDMInRoom = False
	Public Activate_PSDMInRoom As Boolean
	
	'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Public Const Max_int_ROOM_NCOINI As Short = 400 'THIS MAXIMUM IS ALSO LOCATED IN THE FORTRAN MODULE.
	Public Const Max_int_ROOM_NEMITINI As Short = 400 'THIS MAXIMUM IS ALSO LOCATED IN THE FORTRAN MODULE.
	Public Const Max_int_ROOM_NKINI As Short = 400 'THIS MAXIMUM IS ALSO LOCATED IN THE FORTRAN MODULE.
	Structure RoomParam_Type
		'---- INPUT ROOM PARAMETERS: ----
		Dim COUNT_CONTAMINANT As Short
		Dim ROOM_VOL As Double 'm^3
		Dim ROOM_FLOWRATE As Double 'm^3/s
		<VBFixedArray(Number_Compo_Max)> Dim ROOM_C0() As Double 'mg/L
		<VBFixedArray(Number_Compo_Max)> Dim ROOM_EMIT() As Double 'ug/s
		'---- CALCULATED ROOM PARAMETERS: ----
		Dim ROOM_CHANGE_RATE As Double 'hour^(-1)
		<VBFixedArray(Number_Compo_Max)> Dim ROOM_SS_VALUE() As Double 'ug/L
		'---- UNITS FOR ALL VARIABLES: ----
		Dim ROOM_VOL_Units As String 'default: m^3
		Dim ROOM_FLOWRATE_Units As String 'default: m^3/s
		Dim ROOM_C0_Units As String 'default: mg/L
		Dim ROOM_EMIT_Units As String 'default: ug/s
		Dim INITIAL_ROOM_CONC_Units As String 'default: mg/L
		'---- NEW AS OF 9/16/98: ----
		<VBFixedArray(Number_Compo_Max)> Dim INITIAL_ROOM_CONC() As Double 'mg/L
		'---- NEW AS OF 9/16/98 ENDS. ----
		'---- NEW AS OF 8/18/99: ----
		<VBFixedArray(Number_Compo_Max)> Dim RXN_RATE_CONSTANT() As Double
		'(i): first-order rate constant for destruction of chemical i, 1/s
		<VBFixedArray(Number_Compo_Max)> Dim RXN_PRODUCT() As Short
		'(i): index of chemical that is produced through destruction of chemical i (or 0 if none), unitless
		<VBFixedArray(Number_Compo_Max)> Dim RXN_RATIO() As Double
		'(i): number of moles of chemical RXN_PRODUCT(i) produced by destruction of 1 mole of chemical i
		'---- NEW AS OF 8/18/99 ENDS. ----
		'---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
		'
		'/////////   TIME-VARIABLE Co   //////////////////////////////////
		''''bool_ROOM_COINI_ISTIMEVAR As Boolean
		''''int_ROOM_NCOINI As Integer
		''''dbl_ROOM_TCOINI() As Double   '(x): x=row
		Dim bool_ROOM_COINI_ISTIMEVAR() As Boolean '(x): x=chemical
		Dim int_ROOM_NCOINI() As Short '(x): x=chemical
		Dim dbl_ROOM_TCOINI(,) As Double '(x,y): x=chemical, y=row   (minutes)
		Dim dbl_ROOM_COINI(,) As Double '(x,y): x=chemical, y=row   (ug/L)
		Dim u_ROOM_TCOINI As String 'units of display
		Dim u_ROOM_COINI As String 'units of display
		'
		'/////////   TIME-VARIABLE w*A   /////////////////////////////////
		Dim bool_ROOM_EMITINI_ISTIMEVAR() As Boolean '(x): x=chemical
		Dim int_ROOM_NEMITINI() As Short '(x): x=chemical
		Dim dbl_ROOM_TEMITINI(,) As Double '(x,y): x=chemical, y=row   (minutes)
		Dim dbl_ROOM_EMITINI(,) As Double '(x,y): x=chemical, y=row   (ug/s)
		Dim u_ROOM_TEMITINI As String 'units of display
		Dim u_ROOM_EMITINI As String 'units of display
		'---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
		'---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
		'
		'/////////   TIME-VARIABLE K   /////////////////////////////////
		Dim bool_ROOM_KINI_ISTIMEVAR() As Boolean '(x): x=chemical
		Dim int_ROOM_NKINI() As Short '(x): x=chemical
		Dim dbl_ROOM_TKINI(,) As Double '(x,y): x=chemical, y=row   (minutes)
		Dim dbl_ROOM_KINI(,) As Double '(x,y): x=chemical, y=row   (ug/s)
		Dim u_ROOM_TKINI As String 'units of display
		Dim u_ROOM_KINI As String 'units of display
		'---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array ROOM_C0 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If ROOM_C0 Is Nothing Then ReDim ROOM_C0(Number_Compo_Max)
			'UPGRADE_WARNING: Lower bound of array ROOM_EMIT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If ROOM_EMIT Is Nothing Then ReDim ROOM_EMIT(Number_Compo_Max)
			'UPGRADE_WARNING: Lower bound of array ROOM_SS_VALUE was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If ROOM_SS_VALUE Is Nothing Then ReDim ROOM_SS_VALUE(Number_Compo_Max)
			'UPGRADE_WARNING: Lower bound of array INITIAL_ROOM_CONC was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If INITIAL_ROOM_CONC Is Nothing Then ReDim INITIAL_ROOM_CONC(Number_Compo_Max)
			'UPGRADE_WARNING: Lower bound of array RXN_RATE_CONSTANT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If RXN_RATE_CONSTANT Is Nothing Then ReDim RXN_RATE_CONSTANT(Number_Compo_Max)
			'UPGRADE_WARNING: Lower bound of array RXN_PRODUCT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If RXN_PRODUCT Is Nothing Then ReDim RXN_PRODUCT(Number_Compo_Max)
			'UPGRADE_WARNING: Lower bound of array RXN_RATIO was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If RXN_RATIO Is Nothing Then ReDim RXN_RATIO(Number_Compo_Max)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure RoomParams may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public RoomParams As RoomParam_Type
	'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
	
	
	
	
	
	
	Const Structs_PSDMInRoom_declarations_end As Boolean = True
	
	
	Sub RoomParam_Recalculate(ByRef rp As RoomParam_Type)
		Dim i As Short
		'(ROOM_CHANGE_RATE,hour^-1) =
		'(ROOM_FLOWRATE,m^3/s)/(ROOM_VOL,m^3)*(60 s/min)*(60 min/hr)
		rp.ROOM_CHANGE_RATE = rp.ROOM_FLOWRATE / rp.ROOM_VOL * 60# * 60#
		For i = 1 To Number_Compo_Max
			'(ROOM_SS_VALUE,ug/L) =
			'(ROOM_C0,mg/L)*(1000 ug/mg) +
			'(ROOM_EMIT,ug/s)/((ROOM_FLOWRATE,m^3/s)*(1000 L/m^3))
			rp.ROOM_SS_VALUE(i) = rp.ROOM_C0(i) * 1000# + rp.ROOM_EMIT(i) / (rp.ROOM_FLOWRATE * 1000#)
		Next i
	End Sub
	
	
	
	
	
	'Option Explicit
	'
	''
	'' NOTE:
	'' IF Distribute_PSDMInRoom IS SET TO FALSE, THEN
	'' Activate_PSDMInRoom IS SET TO FALSE.
	'' IF Distribute_PSDMInRoom IS SET TO TRUE, THEN
	'' THE EXISTENCE OF THE FILE $(AppPath)\PSDMROOM.DAT IS
	'' CHECKED.  IF THIS FILE EXISTS, Activate_PSDMInRoom
	'' IS SET TO TRUE.  IF THIS FILE DOES NOT EXIST,
	'' Activate_PSDMInRoom IS SET TO FALSE.
	''
	'' IF Activate_PSDMInRoom IS SET TO TRUE, THE PSDMR
	'' MENU ENTRIES AND FILE-LOAD CAPABILITIES ARE ENABLED;
	'' IF SET TO FALSE, THESE CAPABILITIES ARE DISABLED.
	''
	'Global Const Distribute_PSDMInRoom = True
	''Global Const Distribute_PSDMInRoom = False
	'Global Activate_PSDMInRoom As Boolean
	'
	''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'Global Const Max_int_ROOM_NCOINI = 400      'THIS MAXIMUM IS ALSO LOCATED IN THE FORTRAN MODULE.
	'Type RoomParam_Type
	'  '---- INPUT ROOM PARAMETERS: ----
	'  COUNT_CONTAMINANT As Integer
	'  ROOM_VOL As Double                          'm^3
	'  ROOM_FLOWRATE As Double                     'm^3/s
	'  ROOM_C0(1 To Number_Compo_Max) As Double    'mg/L
	'  ROOM_EMIT(1 To Number_Compo_Max) As Double  'ug/s
	'  '---- CALCULATED ROOM PARAMETERS: ----
	'  ROOM_CHANGE_RATE As Double                  'hour^(-1)
	'  ROOM_SS_VALUE(1 To Number_Compo_Max) As Double  'ug/L
	'  '---- UNITS FOR ALL VARIABLES: ----
	'  ROOM_VOL_Units As String        'default: m^3
	'  ROOM_FLOWRATE_Units As String   'default: m^3/s
	'  ROOM_C0_Units As String         'default: mg/L
	'  ROOM_EMIT_Units As String       'default: ug/s
	'  INITIAL_ROOM_CONC_Units As String  'default: mg/L
	'  '---- NEW AS OF 9/16/98: ----
	'  INITIAL_ROOM_CONC(1 To Number_Compo_Max) As Double 'mg/L
	'  '---- NEW AS OF 9/16/98 ENDS. ----
	'  '---- NEW AS OF 8/18/99: ----
	'  RXN_RATE_CONSTANT(1 To Number_Compo_Max) As Double
	'      '(i): first-order rate constant for destruction of chemical i, 1/s
	'  RXN_PRODUCT(1 To Number_Compo_Max) As Integer
	'      '(i): index of chemical that is produced through destruction of chemical i (or 0 if none), unitless
	'  RXN_RATIO(1 To Number_Compo_Max) As Double
	'      '(i): number of moles of chemical RXN_PRODUCT(i) produced by destruction of 1 mole of chemical i
	'  '---- NEW AS OF 8/18/99 ENDS. ----
	'  '---- NEW AS OF 11/11/99 BEGINS: ----
	'  bool_ROOM_COINI_ISTIMEVAR As Boolean
	'  int_ROOM_NCOINI As Integer
	'  dbl_ROOM_TCOINI() As Double   '(x): x=row
	'  dbl_ROOM_COINI() As Double    '(x,y): x=chemical, y=row
	'  u_ROOM_TCOINI As String
	'  u_ROOM_COINI As String
	'
	'
	'  bool_ROOM_EMITINI_ISTIMEVAR As Boolean
	'
	'
	'  '---- NEW AS OF 11/11/99 ENDS. ----
	'End Type
	'Global RoomParams As RoomParam_Type
	''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'
	'
	'
	'
	'
	'
	'
	'Const Structs_PSDMInRoom_declarations_end = True
	'
	'
	'Sub RoomParam_Recalculate(rp As RoomParam_Type)
	'Dim i As Integer
	'  '(ROOM_CHANGE_RATE,hour^-1) =
	'  '(ROOM_FLOWRATE,m^3/s)/(ROOM_VOL,m^3)*(60 s/min)*(60 min/hr)
	'  rp.ROOM_CHANGE_RATE = rp.ROOM_FLOWRATE / rp.ROOM_VOL * 60# * 60#
	'  For i = 1 To Number_Compo_Max
	'    '(ROOM_SS_VALUE,ug/L) =
	'    '(ROOM_C0,mg/L)*(1000 ug/mg) +
	'    '(ROOM_EMIT,ug/s)/((ROOM_FLOWRATE,m^3/s)*(1000 L/m^3))
	'    rp.ROOM_SS_VALUE(i) = rp.ROOM_C0(i) * 1000# + rp.ROOM_EMIT(i) / (rp.ROOM_FLOWRATE * 1000#)
	'  Next i
	'End Sub
	'
	'
	'
End Module