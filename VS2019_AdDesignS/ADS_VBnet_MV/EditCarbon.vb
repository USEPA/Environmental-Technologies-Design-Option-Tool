Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmEditCarbon
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer

	Dim FORM_MODE As Short
	Const FORM_MODE_QUERY_DATABASE As Short = 1
	Const FORM_MODE_EDIT_DATABASE As Short = 2
	
	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_USE_THESE As Boolean

	'UPGRADE_ISSUE: Database object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Dim DB_Carbon As DAO.Database


	Private LocalCarbon_Record As frmEditCarbonData_Record_Type
	
	
	
	
	Const frmEditCarbon_declarations_end As Boolean = True
	
	
	Sub frmEditCarbon_QueryDatabase(ByRef OUTPUT_User_Transferred_Data As Boolean)
		
		On Error GoTo err_frmEditCarbon_QueryDatabase
		'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
		'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
		'UPGRADE_WARNING: Couldn't resolve default property of object Ws1.OpenDatabase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'DB_Carbon = DAOEngine.OpenDatabase(fn_DB_Carbon)  'From ws1 to Daoengine ?? Shang
		'DB_Carbon = Ws1.OpenDatabase(fn_DB_Carbon) 'Throws error
		'replace w/ password
		DB_Carbon = Ws1.OpenDatabase(fn_DB_Carbon, True, False, ";pwd=" & decrypt_string(Encrypted_User_Password))
		FORM_MODE = FORM_MODE_QUERY_DATABASE
		Me.ShowDialog()
		If (USER_HIT_USE_THESE) Then
			OUTPUT_User_Transferred_Data = True
		Else
			OUTPUT_User_Transferred_Data = False
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DB_Carbon.Close()
		Exit Sub
exit_err_frmEditCarbon_QueryDatabase: 
		Exit Sub
err_frmEditCarbon_QueryDatabase: 
		Call Show_Trapped_Error("frmEditCarbon_QueryDatabase")
		OUTPUT_User_Transferred_Data = False
		Resume exit_err_frmEditCarbon_QueryDatabase
	End Sub
	Sub frmEditCarbon_EditDatabase()
		On Error GoTo err_frmEditCarbon_EditDatabase
		'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
		'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
		'UPGRADE_WARNING: Couldn't resolve default property of object Ws1.OpenDatabase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'DB_Carbon = DAOEngine.OpenDatabase(fn_DB_Carbon)   ' ws1 to DaoEngine
		DB_Carbon = Ws1.OpenDatabase(fn_DB_Carbon, True, False, ";pwd=" & decrypt_string(Encrypted_User_Password))
		'Set DB_Carbon = ws1.OpenDatabase(fn_DB_Carbon)
		FORM_MODE = FORM_MODE_EDIT_DATABASE
		Me.ShowDialog()
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DB_Carbon.Close()
		Exit Sub
exit_err_frmEditCarbon_EditDatabase: 
		Exit Sub
err_frmEditCarbon_EditDatabase: 
		Call Show_Trapped_Error("frmEditCarbon_EditDatabase")
		Resume exit_err_frmEditCarbon_EditDatabase
	End Sub
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdOK.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdOK.Enabled = False
			mnuManufacturerItem(1).Enabled = False
			mnuManufacturerItem(2).Enabled = False
			mnuManufacturerItem(3).Enabled = False
			mnuAdsorbentItem(1).Enabled = False
			mnuAdsorbentItem(2).Enabled = False
			mnuAdsorbentItem(3).Enabled = False
		End If
	End Sub
	
	
	Sub populate_lstManu()
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim Current_Criteria As String
		Dim SAVE_CURRENT_POSITION As Integer
		Dim NEW_LISTINDEX As Short
		Dim This_ID As Integer
		Dim NumRecords As Integer
		On Error GoTo err_populate_lstManu
		'SAVE CURRENT POSITION.
		If (lstManu.Items.Count > 0) And (lstManu.SelectedIndex >= 0) Then
			SAVE_CURRENT_POSITION = VB6.GetItemData(lstManu, lstManu.SelectedIndex)
		Else
			SAVE_CURRENT_POSITION = -1
		End If
		'SET UP SEARCH CRITERIA.
		Current_Criteria = "select * from [Manufacturers] " & "order by [Name]"
		'START SEARCH.
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumRecords = Rs1.RecordCount
		On Error GoTo err_populate_lstManu
		'POPULATE LISTBOX.
		lstManu.Items.Clear()
		If (NumRecords = 0) Then
			'NO RECORDS AVAILABLE.
			lstManu.Visible = False
			lblEmpty_lstManu.SetBounds(lstManu.Left, lstManu.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			lblEmpty_lstManu.Visible = True
		Else
			'DISPLAY RECORDS.
			lstManu.Visible = True
			lblEmpty_lstManu.Visible = False
			NEW_LISTINDEX = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Do Until Rs1.EOF
				'UPGRADE_WARNING: Untranslated statement in populate_lstManu. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstManu. Please check source code.
				'UPGRADE_ISSUE: ListBox property lstManu.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'

				'lstManu.AddItem Database_Get_String(Rs1, "Name")
				NEW_LISTINDEX = lstManu.Items.Add(Database_Get_String(Rs1, "Name"))

				This_ID = Database_Get_Long(Rs1, "Manufacturer ID")

				'lstManu.Items(NEW_LISTINDEX).itemdata = This_ID
				VB6.SetItemData(lstManu, NEW_LISTINDEX, This_ID)
				If (SAVE_CURRENT_POSITION <> -1) Then
					If (SAVE_CURRENT_POSITION = This_ID) Then
						'UPGRADE_ISSUE: ListBox property lstManu.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
						NEW_LISTINDEX = lstManu.SelectedIndex
					End If
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveNext()
			Loop 
			If (lstManu.Items.Count > 0) Then
				lstManu.SelectedIndex = 0 'NEW_LISTINDEX
			End If
		End If
		'CLOSE DATABASE AND EXIT.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		Exit Sub
exit_err_populate_lstManu: 
		Exit Sub
err_populate_lstManu: 
		Call Show_Trapped_Error("populate_lstManu")
		Resume exit_err_populate_lstManu
	End Sub
	Sub populate_lstName(ByRef THIS_ITEMDATA As Integer)
		Dim PHASE_CODE As Short
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        Dim Rs1 As dao.Recordset
		Dim Current_Criteria As String
		Dim SAVE_CURRENT_POSITION As Integer
		Dim This_ID As Integer
		Dim NEW_LISTINDEX As Integer
		Dim NumRecords As Integer
		Dim New_Index As Integer
		On Error GoTo err_populate_lstName
		Dim SavedIndex = lstName.SelectedIndex

		'GET PHASE CODE.
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optPhase_0.Checked) Then PHASE_CODE = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optPhase_1.Checked) Then PHASE_CODE = 2
		'SAVE CURRENT POSITION.
		If (lstName.Items.Count > 0) And (lstName.SelectedIndex >= 0) Then
			SAVE_CURRENT_POSITION = VB6.GetItemData(lstName, lstName.SelectedIndex)
		Else
			SAVE_CURRENT_POSITION = -1
		End If
		'SET UP SEARCH CRITERIA.
		Current_Criteria = "select * from [Carbon Data] " & "where [Manufacturer ID]=" & Trim(Str(THIS_ITEMDATA)) & " and " & "[Phase ID]=" & Trim(Str(PHASE_CODE)) & " " & "order by [Name]"
		'START SEARCH.
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumRecords = Rs1.RecordCount
		On Error GoTo err_populate_lstName
		'POPULATE LISTBOX.

		lstName.Items.Clear()

		If SavedIndex < 0 Then
			SavedIndex = 0
		ElseIf SavedIndex >= NumRecords Then
			SavedIndex = NumRecords - 1
		End If

		If (NumRecords = 0) Then
			'NO RECORDS AVAILABLE.
			lstName.Visible = False
			lblEmpty_lstName.SetBounds(lstName.Left, lstName.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			lblEmpty_lstName.Visible = True
		Else
			'DISPLAY RECORDS.
			lstName.Visible = True
			lblEmpty_lstName.Visible = False
			NEW_LISTINDEX = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Do Until Rs1.EOF
				'UPGRADE_WARNING: Untranslated statement in populate_lstName. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstName. Please check source code.
				'UPGRADE_ISSUE: ListBox property lstName.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'

				New_Index = lstName.Items.Add(Database_Get_String(Rs1, "Name"))
				This_ID = Database_Get_Long(Rs1, "ID")
				VB6.SetItemData(lstName, New_Index, This_ID)
				If (SAVE_CURRENT_POSITION <> -1) Then
					If (SAVE_CURRENT_POSITION = This_ID) Then
						'UPGRADE_ISSUE: ListBox property lstName.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
						NEW_LISTINDEX = lstName.SelectedIndex
					End If
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveNext()
			Loop 
			If (lstName.Items.Count > 0) Then
				lstName.SelectedIndex = NEW_LISTINDEX
			End If
		End If
		'CLOSE DATABASE AND EXIT.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		lstName.SelectedIndex = SavedIndex
		Exit Sub
exit_err_populate_lstName: 
		Exit Sub
err_populate_lstName: 
		Call Show_Trapped_Error("populate_lstName")
		Resume exit_err_populate_lstName
	End Sub
	Sub populate_lblData(ByRef THIS_ITEMDATA As Integer)
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        Dim Rs1 As dao.Recordset
		Dim NumRecords As Integer
		Dim Current_Criteria As String
		Dim TempDbl As Double
		On Error GoTo err_populate_lblData
		'SET UP SEARCH CRITERIA.
		Current_Criteria = "select * from [Carbon Data] " & "where [ID]=" & Trim(Str(THIS_ITEMDATA)) & " " & "order by [Name]"
		'START SEARCH.
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumRecords = Rs1.RecordCount
		On Error GoTo err_populate_lblData
		'POPULATE LABEL CONTROLS.
		If (NumRecords = 0) Then
			'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
			'DO NOTHING.
		Else
			'DISPLAY (FIRST) RECORD (THERE SHOULD ONLY BE ONE).
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.
			LocalCarbon_Record.Name = Database_Get_String(Rs1, "Name")
			LocalCarbon_Record.Manufacturer = ""
			On Error Resume Next
			LocalCarbon_Record.Manufacturer = VB6.GetItemString(lstManu, lstManu.SelectedIndex)
			On Error GoTo err_populate_lblData
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.
			'NOTE: RADIUS IN DATABASE IS STORED IN millimeters;
			'DIVISION BY 10 CONVERTS THIS TO centimers.
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.

			LocalCarbon_Record.AppDen = Database_Get_Double(Rs1, "Apparent Density")
			'NOTE: RADIUS IN DATABASE IS STORED IN millimeters;
			'DIVISION BY 10 CONVERTS THIS TO centimers.
			LocalCarbon_Record.ParticleRadius =
			Database_Get_Double(Rs1, "Average Particle Radius") / 10.0#
			LocalCarbon_Record.ParticlePorosity = Database_Get_Double(Rs1, "Porosity")
			LocalCarbon_Record.AdsType = Database_Get_String(Rs1, "Type")

			Call AssignCaptionAndTag(lblData(0), LocalCarbon_Record.AppDen)
			Call AssignCaptionAndTag(lblData(1), LocalCarbon_Record.ParticleRadius)
			Call AssignCaptionAndTag(lblData(2), LocalCarbon_Record.ParticlePorosity)
			Call AssignCaptionAndTag(lblData(3), LocalCarbon_Record.AdsType)
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.

			TempDbl = Database_Get_Double(Rs1, "W0")
			LocalCarbon_Record.W0 = TempDbl
			If (TempDbl = 0#) Then
				Call AssignCaptionAndTag(lblData(4), "Unavailable")
			Else
				Call AssignCaptionAndTag(lblData(4), TempDbl)
			End If
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.

			TempDbl = Database_Get_Double(Rs1, "BB")
			LocalCarbon_Record.BB = TempDbl
			If (TempDbl = 0#) Then
				Call AssignCaptionAndTag(lblData(5), "Unavailable")
			Else
				Call AssignCaptionAndTag(lblData(5), TempDbl)
			End If
			'UPGRADE_WARNING: Untranslated statement in populate_lblData. Please check source code.

			TempDbl = Database_Get_Double(Rs1, "Polanyi Exponent")
			LocalCarbon_Record.PolanyiExponent = TempDbl
			If (TempDbl = 0#) Then
				Call AssignCaptionAndTag(lblData(6), "Unavailable")
			Else
				Call AssignCaptionAndTag(lblData(6), TempDbl)
			End If
		End If
		'CLOSE DATABASE AND EXIT.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		Exit Sub
exit_err_populate_lblData: 
		Exit Sub
err_populate_lblData: 
		Call Show_Trapped_Error("populate_lblData")
		Resume exit_err_populate_lblData
	End Sub
	
	
	Private Sub cmdCancel_Click()
		'frmEditAdsorber_Cancelled = True
		USER_HIT_CANCEL = True
		USER_HIT_USE_THESE = False
		Me.Dispose()
	End Sub
	Private Sub cmdOK_Click()
		If (lstManu.SelectedIndex < 0) Or (lstManu.Items.Count = 0) Then
			Call Show_Error("You must first select a manufacturer.")
			Exit Sub
		End If
		If (lstName.SelectedIndex < 0) Or (lstName.Items.Count = 0) Then
			Call Show_Error("You must first select an adsorbent.")
			Exit Sub
		End If
		If (LocalCarbon_Record.AppDen = 0#) Or (LocalCarbon_Record.ParticlePorosity = 0#) Or (LocalCarbon_Record.ParticleRadius = 0#) Then
			Call Show_Error("You must select an adsorbent with non-zero values " & "for apparent density, particle porosity, and particle radius.")
			Exit Sub
		End If
		'
		' TRANSFER DATA TO CURRENT CARBON RECORD.
		'
		Carbon.Name = Trim(LocalCarbon_Record.Manufacturer)
		If (Carbon.Name <> "") Then
			Carbon.Name = Carbon.Name & " "
		End If
		Carbon.Name = Carbon.Name & Trim(LocalCarbon_Record.Name)
		Carbon.Density = LocalCarbon_Record.AppDen
		Carbon.ParticleRadius = LocalCarbon_Record.ParticleRadius / 100#
		Carbon.Porosity = LocalCarbon_Record.ParticlePorosity
		Carbon.ShapeFactor = 1#
		Carbon.W0 = LocalCarbon_Record.W0
		Carbon.BB = LocalCarbon_Record.BB
		Carbon.PolanyiExponent = LocalCarbon_Record.PolanyiExponent
		'
		' EXIT OUT OF HERE.
		'
		USER_HIT_CANCEL = False
		USER_HIT_USE_THESE = True
		Me.Dispose()
	End Sub
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub
	
	Private Sub frmEditCarbon_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'MISC INITS.
		'		Me.Height = VB6.TwipsToPixelsY(7290)
		'		Me.Width = VB6.TwipsToPixelsX(9600)
		Call CenterOnForm(Me, frmMain)
		lblUnit(0).Text = "g/cm³"
		lblUnit(4).Text = "cm³/g"
		If (FORM_MODE = FORM_MODE_QUERY_DATABASE) Then
			'QUERY DATABASE MODE.
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdOK.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdOK.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancel.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCancel.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancel.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCancel.Text = "&Cancel"
		Else
			'EDIT DATABASE MODE.
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdOK.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdOK.Visible = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancel.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCancel.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancel.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCancel.Text = "E&xit"
		End If
		If (Bed.Phase = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optPhase_0.Checked = True
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optPhase_1.Checked = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optPhase_0.Checked = False
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optPhase_1.Checked = True
		End If
		'RE-POPULATE MANUFACTURER LIST.
		Call populate_lstManu()
		'DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	
	
	'UPGRADE_WARNING: Event lstManu.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstManu_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstManu.SelectedIndexChanged
		Dim THIS_ITEMDATA As Integer
		If (lstManu.SelectedIndex < 0) Or (lstManu.Items.Count <= 0) Then
			Exit Sub
		End If
		THIS_ITEMDATA = VB6.GetItemData(lstManu, lstManu.SelectedIndex)
		Call populate_lstName(THIS_ITEMDATA)
	End Sub
	Private Sub lstManu_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstManu.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If ((Button And 2) = 2) Then
			'UPGRADE_ISSUE: Form method frmEditCarbon.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'Me.PopupMenu(mnuManufacturer)
		End If
	End Sub
	'UPGRADE_WARNING: Event lstName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstName.SelectedIndexChanged
		Dim THIS_ITEMDATA As Integer
		If (lstName.SelectedIndex < 0) Or (lstName.Items.Count <= 0) Then
			Exit Sub
		End If
		THIS_ITEMDATA = VB6.GetItemData(lstName, lstName.SelectedIndex)
		Call populate_lblData(THIS_ITEMDATA)
	End Sub
	Private Sub lstName_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstName.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If ((Button And 2) = 2) Then
			'UPGRADE_ISSUE: Form method frmEditCarbon.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'Me.PopupMenu(mnuAdsorbent)
		End If
	End Sub
	
	
	Public Sub mnuAdsorbentItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAdsorbentItem.Click
		Dim Index As Short = mnuAdsorbentItem.GetIndex(eventSender)
		Dim USER_HIT_CANCEL As Boolean
		Dim USER_HIT_SAVE As Boolean
		Dim USER_HIT_SAVEASNEW As Boolean
		Dim THIS_MANU_ID As Integer
		Dim THIS_ADS_ID As Integer
		Dim Current_Criteria As String
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim NumRecords As Integer
		Dim DEFAULT_PHASE_IS_LIQUID As Boolean
		Const INDEX_NEW As Short = 1
		Const INDEX_EDIT As Short = 2
		Const INDEX_DELETE As Short = 3
		Dim NewName As String
		Dim msg As String
		Dim RetVal As Short
		Dim Select_Index As Short
		Dim i As Short
		On Error GoTo err_mnuAdsorbentItem_Click
		If (lstManu.SelectedIndex < 0) Or (lstManu.Items.Count = 0) Then
			Call Show_Error("You must first select a manufacturer.")
			Exit Sub
		End If
		THIS_MANU_ID = VB6.GetItemData(lstManu, lstManu.SelectedIndex)
		If (Index = INDEX_EDIT) Or (Index = INDEX_DELETE) Then
			If (lstName.SelectedIndex < 0) Or (lstName.Items.Count = 0) Then
				Call Show_Error("You must first select an adsorbent.")
				Exit Sub
			End If
			THIS_ADS_ID = VB6.GetItemData(lstName, lstName.SelectedIndex)
			'SET UP SEARCH CRITERIA.
			Current_Criteria = "select * from [Carbon Data] " & "where [ID]=" & Trim(Str(THIS_ADS_ID)) & " " & "order by [Name]"
			'START SEARCH.
			'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.MoveFirst()
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.MoveLast()
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.MoveFirst()
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NumRecords = Rs1.RecordCount
			On Error GoTo err_mnuAdsorbentItem_Click
			'POPULATE LABEL CONTROLS.
			If (NumRecords = 0) Then
				'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
				'EXIT SUBROUTINE.
				Exit Sub
			End If
			'Call AssignCaptionAndTag(lblData(0), Database_Get_Double(RS1, "Apparent Density"))
			'Call AssignCaptionAndTag(lblData(1), Database_Get_Double(RS1, "Average Particle Radius"))
			'Call AssignCaptionAndTag(lblData(2), Database_Get_Double(RS1, "Porosity"))
			'Call AssignCaptionAndTag(lblData(3), Database_Get_String(RS1, "Type"))
		End If
		Select Case Index
			Case INDEX_NEW 'NEW ADSORBENT.
				'DETERMINE DEFAULT PHASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (_optPhase_0.Checked) Then
					DEFAULT_PHASE_IS_LIQUID = True
				Else
					DEFAULT_PHASE_IS_LIQUID = False
				End If
				'ALLOW USER TO ADD NEW RECORD.
				Call frmEditCarbonData.frmEditCarbonData_AddNew(DEFAULT_PHASE_IS_LIQUID, USER_HIT_CANCEL, USER_HIT_SAVE)
				If (USER_HIT_CANCEL) Then Exit Sub
				'ADD THE NEW ADSORBENT RECORD.
				Current_Criteria = "select * from [Carbon Data]"
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'SET THE MANUFACTURER-ID AND PHASE FIELDS.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Manufacturer ID").Value = THIS_MANU_ID
				If (frmEditCarbonData_Record.PhaseIsLiquid) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase ID").Value = 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase ID").Value = 2
				End If
				'THE FIELD [ID] IS AUTOMATICALLY UPDATED.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = frmEditCarbonData_Record.Name
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Type").Value = frmEditCarbonData_Record.AdsType
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Apparent Density").Value = frmEditCarbonData_Record.AppDen
				' NEXT LINE CONVERTS centimeters TO millimeters.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Average Particle Radius").Value = frmEditCarbonData_Record.ParticleRadius * 10.0#
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Porosity").Value = frmEditCarbonData_Record.ParticlePorosity
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("W0").Value = frmEditCarbonData_Record.W0
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("BB").Value = frmEditCarbonData_Record.BB
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Polanyi Exponent").Value = frmEditCarbonData_Record.PolanyiExponent
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				THIS_ADS_ID = Database_Get_Long(Rs1, "ID")
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'CLOSE THE DATABASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'UPDATE WINDOW.
				Call lstManu_SelectedIndexChanged(lstManu, New System.EventArgs())
				'SELECT THE NEW ADSORBENT.
				Select_Index = 0
				For i = 0 To lstName.Items.Count - 1
					If (VB6.GetItemData(lstName, i) = THIS_ADS_ID) Then
						Select_Index = i
						Exit For
					End If
				Next i
				If (lstName.Items.Count > 0) Then
					lstName.SelectedIndex = Select_Index
				End If
			Case INDEX_EDIT 'EDIT ADSORBENT.
				'TRANSFER DATABASE RECORD FIELDS TO LOCAL MEMORY.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				' NEXT LINE CONVERTS millimeters TO centimeters.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuAdsorbentItem_Click. Please check source code.


				If (Database_Get_Long(Rs1, "Phase ID") = 1) Then
					frmEditCarbonData_Record.PhaseIsLiquid = True
				Else
					frmEditCarbonData_Record.PhaseIsLiquid = False
				End If
				frmEditCarbonData_Record.AppDen = Database_Get_Double(Rs1, "Apparent Density")
				' NEXT LINE CONVERTS millimeters TO centimeters.
				frmEditCarbonData_Record.ParticleRadius =
				Database_Get_Double(Rs1, "Average Particle Radius") / 10.0#
				frmEditCarbonData_Record.ParticlePorosity = Database_Get_Double(Rs1, "Porosity")
				frmEditCarbonData_Record.W0 = Database_Get_Double(Rs1, "W0")
				frmEditCarbonData_Record.BB = Database_Get_Double(Rs1, "BB")
				frmEditCarbonData_Record.PolanyiExponent = Database_Get_Double(Rs1, "Polanyi Exponent")
				frmEditCarbonData_Record.Name = Database_Get_String(Rs1, "Name")
				frmEditCarbonData_Record.AdsType = Database_Get_String(Rs1, "Type")



				'ALLOW USER TO EDIT THIS RECORD.
				Call frmEditCarbonData.frmEditCarbonData_Edit(USER_HIT_CANCEL, USER_HIT_SAVE, USER_HIT_SAVEASNEW)
				If (USER_HIT_CANCEL) Then Exit Sub
				'SAVE THE EDITED ADSORBENT RECORD.
				Current_Criteria = "select * from [Carbon Data] " & "where [ID]=" & Trim(Str(THIS_ADS_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
				'SET THE MANUFACTURER-ID AND PHASE FIELDS.
				If (USER_HIT_SAVE) Then
					'MODIFY EXISTING RECORD.
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Edit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Edit()
					'KEEP ORIGINAL [Manufacturer ID] FIELD INTACT.
					'KEEP ORIGINAL [ID] FIELD INTACT.
				End If
				If (USER_HIT_SAVEASNEW) Then
					'GENERATE NEW RECORD.
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.AddNew()
					'SAVE [Manufacturer ID] FIELD.
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Manufacturer ID").Value = THIS_MANU_ID
					'THE FIELD [ID] IS AUTOMATICALLY CREATED DURING THE .Update COMMAND.
				End If
				If (frmEditCarbonData_Record.PhaseIsLiquid) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase ID").Value = 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase ID").Value = 2
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = frmEditCarbonData_Record.Name
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Type").Value = frmEditCarbonData_Record.AdsType
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Apparent Density").Value = frmEditCarbonData_Record.AppDen
				' NEXT LINE CONVERTS centimeters TO millimeters.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Average Particle Radius").Value = frmEditCarbonData_Record.ParticleRadius * 10.0#
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Porosity").Value = frmEditCarbonData_Record.ParticlePorosity
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("W0").Value = frmEditCarbonData_Record.W0
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("BB").Value = frmEditCarbonData_Record.BB
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Polanyi Exponent").Value = frmEditCarbonData_Record.PolanyiExponent
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'CLOSE THE DATABASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'UPDATE WINDOW.
				Call lstManu_SelectedIndexChanged(lstManu, New System.EventArgs())
			Case INDEX_DELETE 'DELETE ADSORBENT.
				NewName = Trim(VB6.GetItemString(lstName, lstName.SelectedIndex))
				msg = "Do you really want to delete adsorbent '" & NewName & "' ?"
				RetVal = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, AppName_For_Display_Long)
				If RetVal = MsgBoxResult.No Then Exit Sub
				'PERFORM DELETION.
				Current_Criteria = "select * from [Carbon Data] where " & "[ID] = " & Trim(Str(THIS_ADS_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Delete()
				'CLOSE THE DATABASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'UPDATE WINDOW.
				Call lstManu_SelectedIndexChanged(lstManu, New System.EventArgs())
		End Select
		Exit Sub
exit_err_mnuAdsorbentItem_Click: 
		Exit Sub
err_mnuAdsorbentItem_Click: 
		Call Show_Trapped_Error("mnuAdsorbentItem_Click")
		Resume exit_err_mnuAdsorbentItem_Click
	End Sub
	Public Sub mnuManufacturerItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuManufacturerItem.Click
		Dim Index As Short = mnuManufacturerItem.GetIndex(eventSender)
		Dim THIS_MANU_ID As Integer
		Dim NewName As String
		'Dim USER_HIT_CANCEL As Boolean
		Dim Current_Criteria As String
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim msg As String
		Dim RetVal As Short
		Dim i As Short
		Dim Select_Index As Short
		Dim NumRecords As Short
		On Error GoTo err_mnuManufacturerItem_Click
		If (Index = 2) Or (Index = 3) Then
			If (lstManu.SelectedIndex < 0) Or (lstManu.Items.Count = 0) Then
				Call Show_Error("You must first select a manufacturer.")
				Exit Sub
			End If
			THIS_MANU_ID = VB6.GetItemData(lstManu, lstManu.SelectedIndex)
		End If
		Select Case Index
			Case 1 'new
				NewName = "New Manufacturer"
				Do While (1 = 1)
					NewName = frmNewName.frmNewName_GetName("Creating New Manufacturer", "Each manufacturer record should have a unique name.", NewName, USER_HIT_CANCEL)
					If (USER_HIT_CANCEL) Then Exit Sub
					NewName = Trim(NewName)
					If (NewName <> "") Then Exit Do
					Call Show_Error("Manufacturer name must be a non-blank string.")
				Loop 
				'ADD THE NEW MANUFACTURER RECORD.
				Current_Criteria = "select * from [Manufacturers]"
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'THE FIELD [Manufacturer ID] IS AUTOMATICALLY UPDATED.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = NewName
				'UPGRADE_WARNING: Untranslated statement in mnuManufacturerItem_Click. Please check source code.
				THIS_MANU_ID = Database_Get_Long(Rs1, "Manufacturer ID")
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'REDISPLAY WINDOW.
				Call populate_lstManu()
				'SELECT THE NEW MANUFACTURER.
				Select_Index = 0
				For i = 0 To lstManu.Items.Count - 1
					If (VB6.GetItemData(lstManu, i) = THIS_MANU_ID) Then
						Select_Index = i
						Exit For
					End If
				Next i
				If (lstManu.Items.Count > 0) Then
					lstManu.SelectedIndex = Select_Index
				End If
			Case 2 'edit current
				NewName = Trim(VB6.GetItemString(lstManu, lstManu.SelectedIndex))
				Do While (1 = 1)
					NewName = frmNewName.frmNewName_GetName("Editing Existing Manufacturer Name", "Each manufacturer record should have a unique name.", NewName, USER_HIT_CANCEL)
					If (USER_HIT_CANCEL) Then Exit Sub
					NewName = Trim(NewName)
					If (NewName <> "") Then Exit Do
					Call Show_Error("Manufacturer name must be a non-blank string.")
				Loop 
				'EDIT THE MANUFACTURER RECORD.
				Current_Criteria = "select * from [Manufacturers] where " & "[Manufacturer ID] = " & Trim(Str(THIS_MANU_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Edit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Edit()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = NewName
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'REDISPLAY WINDOW.
				Call populate_lstManu()
			Case 3 'delete current
				NewName = Trim(VB6.GetItemString(lstManu, lstManu.SelectedIndex))
				msg = "Do you really want to delete manufacturer '" & NewName & "' and all of the corresponding adsorbent " & "records from the database ?"
				RetVal = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, AppName_For_Display_Long)
				If RetVal = MsgBoxResult.No Then Exit Sub
				'PERFORM DELETION OF MANUFACTURER RECORD.
				Current_Criteria = "select * from [Manufacturers] where " & "[Manufacturer ID] = " & Trim(Str(THIS_MANU_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Delete()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'PERFORM DELETION OF ADSORBENT RECORDS.
				Current_Criteria = "select * from [Carbon Data] where " & "[Manufacturer ID] = " & Trim(Str(THIS_MANU_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Carbon.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Carbon.OpenRecordset(Current_Criteria)
				On Error Resume Next
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveFirst()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveLast()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveFirst()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumRecords = Rs1.RecordCount
				On Error GoTo err_mnuManufacturerItem_Click
				If (NumRecords > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Do Until Rs1.EOF
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Rs1.Delete()
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Rs1.MoveNext()
					Loop 
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'REDISPLAY WINDOW.
				Call populate_lstManu()
				'DISPLAY TOTAL.
				Call Show_Message("A total of " & Trim(Str(NumRecords)) & " adsorbent records were deleted.")
		End Select
		Exit Sub
exit_err_mnuManufacturerItem_Click: 
		Exit Sub
err_mnuManufacturerItem_Click: 
		Call Show_Trapped_Error("mnuManufacturerItem_Click")
		Resume exit_err_mnuManufacturerItem_Click
	End Sub
	
	
	Private Sub optPhase_Click(ByRef Index As Short, ByRef Value As Short)
		Call lstManu_SelectedIndexChanged(lstManu, New System.EventArgs())
	End Sub

	Private Sub cmdOK_ClickEvent(sender As Object, e As EventArgs)
		Call cmdOK_Click()
	End Sub

	Private Sub cmdCancel_ClickEvent(sender As Object, e As EventArgs)
		Call cmdCancel_Click()
	End Sub

	Private Sub _optPhase_0_CheckedChanged(sender As Object, e As EventArgs) Handles _optPhase_0.CheckedChanged
		optPhase_Click(0, 0)
	End Sub

	Private Sub _optPhase_1_CheckedChanged(sender As Object, e As EventArgs) Handles _optPhase_1.CheckedChanged
		optPhase_Click(1, 0)
	End Sub

	Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
		Call cmdCancel_Click()
	End Sub

	Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
		Call cmdOK_Click()
	End Sub

	Private Sub frmEditCarbon_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)
	End Sub
End Class