Option Strict Off
Option Explicit On
Module FileIO
	
	Public Project_Is_Dirty As Boolean


	Const FileIO_declarations_end As Boolean = True
	
	
	'NOTE: THIS FUNCTION WORKS EQUALLY WELL ON
	'EITHER FILES OR DIRECTORIES.
	Function File_IsExists(ByRef fn As String) As Boolean
		Dim Dummy As Integer
		On Error GoTo err_File_IsExists
		Dummy = GetAttr(fn) 'TRIGGERS ERROR IF FILE DOES NOT EXIST.
		File_IsExists = True
		Exit Function
exit_err_File_IsExists: 
		File_IsExists = False
		Exit Function
err_File_IsExists: 
		Resume exit_err_File_IsExists
	End Function
	Function FileExists(ByRef fn As String) As Boolean
		FileExists = File_IsExists(fn)
	End Function

	Function Database_IsTableExist(ByRef Db1 As dao.Database, ByRef Use_TableName As String) As Boolean
		Dim tbl As dao.TableDef
		Dim TableExists As Boolean = False
		For Each tbl In Db1.TableDefs
			If tbl.Name = Use_TableName Then
				TableExists = True
				Exit For
			End If
		Next
		Return TableExists
	End Function

	Sub KillFile_If_Exists(ByRef fn As String)
		If (File_IsExists(fn)) Then
			On Error Resume Next
			Kill(fn)
		End If
	End Sub
	
	
	Sub file_new()
		'DISABLE RESULTS MENU: PSDM, CPHSDM, ECM, COMPARE PSDM, COMPARE CPHSDM.
		frmMain.mnuResultsItem(0).Enabled = False
		frmMain.mnuResultsItem(1).Enabled = False
		frmMain.mnuResultsItem(2).Enabled = False
		frmMain.mnuResultsItem(3).Enabled = False
		frmMain.mnuResultsItem(4).Enabled = False
		''''frmMain.mnuResultsItem(10).Enabled = False      'PSDMR-IN-ROOM.
		'DISABLE OPTIONS MENU: FOULING, INFLUENT CONC, EFFLUENT CONC.
		frmMain.mnuOptionsItem(0).Enabled = False
		frmMain.mnuOptionsItem(1).Enabled = False
		frmMain.mnuOptionsItem(2).Enabled = False
		'DISABLE RUN MENU: PSDM, CPHSDM, ECM.
		frmMain.mnuRunItem(0).Enabled = False
		frmMain.mnuRunItem(1).Enabled = False
		frmMain.mnuRunItem(2).Enabled = False
		frmMain.mnuRunItem(10).Enabled = False 'PSDMR-IN-ROOM.
		frmMain.mnuRunItem(20).Enabled = False 'PSDMR ALONE.
		'DISABLE FILE MENU: SAVE.
		frmMain.mnuFileItem(2).Enabled = False
		'
		' DEMO SETTINGS.
		'
		Call frmMain.frmMain_Reset_DemoVersionDisablings()
		'
		' INITIALIZE FOR LIQUID PHASE DEFAULTS.
		'
		Bed.Phase = 0
		Call Initialize_All_Data(0)
		frmMain.OpenFileDialog1.FileName = ""
		frmMain.Text = AppName_For_Display_Short & "  -  (Untitled)"
		'
		' CLEAR DIRTY (CHANGES) FLAG.
		'

		Project_Is_Dirty = False
		Call DirtyStatus_Set_Current()
	End Sub
	
	
	Sub File_Open(ByRef fn_Open As String)
		'	Dim OpenDatabase As Object
		Dim f As Short
		Dim ThisVersion As Double
		Dim ShowLegacyWarning As Boolean
		Dim OpenedOkay As Boolean
		Dim IsLegacyVersion As Boolean
		'UPGRADE_ISSUE: Database object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim DbTest1 As dao.Database
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim IsInvalidFormat As Boolean
		Dim DataVersion_Major As Short
		Dim DataVersion_Minor As Short
		Dim IsOkay As Boolean
		Dim fnThis As String
		Dim strTest As String
		On Error GoTo err_file_open
		If (IsThisADemo() = True) Then
			IsOkay = False
			fnThis = Trim(UCase(fn_Open))
			strTest = "LIQUID.DAT"
			If Right(fnThis, Len(strTest)) = strTest Then IsOkay = True
			strTest = "GAS.DAT"
			If Right(fnThis, Len(strTest)) = strTest Then IsOkay = True
			If (IsOkay = False) Then
				Call Demo_ShowError("In the demonstration version, only the example files may be opened.")
				''''File_Open = False
				Exit Sub
			End If
		End If
		If (Not FileExists(fn_Open)) Then
			Call Show_Error("File `" & fn_Open & "` does not exist.")
			GoTo exit_sub
		End If
		frmMain.Cursor = System.Windows.Forms.Cursors.WaitCursor
		'DETERMINE WHETHER THIS IS A LEGACY VERSION.
		IsLegacyVersion = True
		On Error Resume Next
		DbTest1 = DAOEngine.OpenDatabase(fn_Open)    'From Ws1 to Daoengine ??? Shang
		If (Err.Number = 0) Then
			IsLegacyVersion = False
			'UPGRADE_WARNING: Couldn't resolve default property of object DbTest1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DbTest1.Close()
		End If
		On Error GoTo err_file_open
		If (IsLegacyVersion = False) Then
			'DETERMINE WHETHER THE MDB FORMAT FILE IS A LEGACY-VERSION.
			IsInvalidFormat = True
			DbTest1 = DAOEngine.OpenDatabase(fn_Open)   'From Ws1 to DaoEngine ??? Shang

			'UPGRADE_WARNING: Untranslated statement in File_Open. Please check source code.
			If (Database_IsTableExist(DbTest1, "Version") = False) Then
				'INVALID FORMAT.
			Else
				Rs1 = DbTest1.OpenRecordset("Version")

				If (Rs1.RecordCount = 0) Then
					'INVALID FORMAT.
				Else
					DataVersion_Major = 0
					DataVersion_Minor = 0
					Rs1.MoveFirst()
					Do Until Rs1.EOF
						Select Case Trim$(UCase$(Rs1("FieldName").Value))
							Case Trim$(UCase$("DataVersion_Major"))
								'DataVersion_Major = Rs1("lngValue").Value
								Call Database_LoadProperty(Rs1, DataVersion_Major)
							Case Trim$(UCase$("DataVersion_Minor"))
								'DataVersion_Minor = Rs1("lngValue").Value
								Call Database_LoadProperty(Rs1, DataVersion_Minor)
								'Case Trim$(UCase$("ContainsPSDMInRoomData"))
								'	Call Database_LoadProperty(Rs1, ContainsPSDMInRoomData)

						End Select
						Rs1.MoveNext()
					Loop
					Rs1.Close()
					DbTest1.Close()
				End If
			End If

			If (DataVersion_Major = 1) Then
				Select Case DataVersion_Minor
					Case 60
						'OPEN A NON-LEGACY-VERSION FILE.
						Call file_new()
						OpenedOkay = File_Open_Latest_v1_60(fn_Open)
						If (OpenedOkay) Then IsInvalidFormat = False
				End Select
			End If
			If (IsInvalidFormat) Then
				Call Show_Error("The selected file is not a " & "valid file.")
				GoTo exit_sub
			End If
		Else
			'OPEN A LEGACY-VERSION FILE (TEXT FORMAT).
			f = FreeFile()
			FileOpen(f, fn_Open, OpenMode.Input)
			Input(f, ThisVersion)
			ShowLegacyWarning = True
			OpenedOkay = True
			Select Case ThisVersion
				Case 1# : Call File_Open_Legacy_v1_00(f)
				Case 1.2 : Call File_Open_Legacy_v1_20(f)
				Case 1.3 : Call File_Open_Legacy_v1_30(f)
				Case 1.42
					If (Activate_PSDMInRoom) Then
						' DO NOTHING.
					Else
						Call Show_Error("Warning: The selected file contains " & "PSDMR model data which is not accessible " & "using this version of the software.  " & "The file will be loaded anyway.")
					End If
					Call File_Open_Legacy_v1_42(f)
				Case 1.4
					''''ShowLegacyWarning = False
					OpenedOkay = File_Open_Legacy_v1_40(f)
				Case Else
					Call Show_Error("The selected file is not a " & "valid file.")
					FileClose(f)
					GoTo exit_sub
			End Select
			If (OpenedOkay = False) Then
				FileClose(f)
				Call file_new()
				GoTo exit_sub
			End If
		End If
		'UPDATE CURRENT BED PHASE.
		Call chem_phase(Bed.Phase)
		'SHOW LEGACY WARNING IF NECESSARY.
		If (ShowLegacyWarning) Then
			'Call Show_Message("Warning: This file is formatted as an " & _
			'"AdXDesignS Version " & Format$(ThisVersion, "0.00") & _
			'" file.  If saved, it will be saved as an AdXDesignS " & _
			'"Version " & Format$(NVersion, "0.00") & _
			'" file.")
			Call Show_Message00("Warning: This file is formatted as a " & "Version " & VB6.Format(ThisVersion, "0.00") & " file.  If saved, it will be saved as a " & "Version " & Trim(Str(Latest_DataVersion_Major)) & "." & Trim(Str(Latest_DataVersion_Minor)) & " file.", MsgBoxStyle.Information, AppName_For_Display_Short & " : Legacy File Version Warning")
		End If
		FileClose(f)
		'UPDATE DISPLAY.
		Filename = fn_Open
		frmMain.Text = AppName_For_Display_Short & "  -  " & Trim(Filename)
		Call frmMain_Refresh()
		'CLEAR DIRTY FLAG.
		Project_Is_Dirty = False
		Call DirtyStatus_Set_Current()
		'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
		Call OldFileList_Promote(Filename, 1, frmMain._mnuFileItem_199, frmMain.mnuFileItem(191), frmMain.mnuFileItem(192), frmMain.mnuFileItem(193), frmMain.mnuFileItem(194))
		GoTo exit_sub
exit_sub: 
		frmMain.Cursor = System.Windows.Forms.Cursors.Default
		Exit Sub
exit_err_file_open: 
		Call file_new()
		GoTo exit_sub
err_file_open:
		'Call Show_Trapped_Error("file_open")
		On Error Resume Next
		FileClose(f)
		Resume exit_err_file_open
	End Sub
	Sub File_OpenAs(ByRef fn_force As String)
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNFileMustExist As Object
		Dim fn_openas As String
		If (fn_force <> "") Then
			fn_openas = fn_force
		Else
			'INPUT NEW FILENAME.
			On Error GoTo err_file_openas
			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.CommonDialog1.DialogTitle = "Open " & AppName_For_Display_Short & " File"

			frmMain.OpenFileDialog1.Title = "Open " & AppName_For_Display_Short & " File"


			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.CommonDialog1.Filter = "All Files (*.*)|*.*|" & AppName_For_Display_Short & " Files (*.dat)|*.dat"

			frmMain.OpenFileDialog1.Filter = "All Files (*.*)|*.*|" & AppName_For_Display_Short & " Files (*.dat)|*.dat"


			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.CommonDialog1.FilterIndex = 2

			frmMain.OpenFileDialog1.FilterIndex = 2

			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.CommonDialog1.CancelError = True

			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNFileMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.CommonDialog1.flags = cdlOFNFileMustExist + cdlOFNPathMustExist

			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.ShowOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.CommonDialog1.ShowOpen()

			frmMain.OpenFileDialog1.ShowDialog()

			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fn_openas = Trim(frmMain.OpenFileDialog1.FileName)
			If (fn_openas = "") Then
				'DO NOTHING.
				Exit Sub
			End If
		End If
		'OPEN THIS FILE.
		Call File_Open(fn_openas)
exit_err_file_openas: 
		Exit Sub
err_file_openas:
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = 75) Then
			'CANCEL BUTTON WAS SELECTED.
			Resume exit_err_file_openas
		End If
		Resume exit_err_file_openas
	End Sub
	
	
	'RETURNS:
	'- true = save went okay.
	'- false = save failed.
	Function File_Save(ByRef fn_Save As String) As Boolean
		Dim f As Short
		Dim SavedOkay As Boolean
		If (IsThisADemo() = True) Then
			Call Demo_ShowError("Saving is not allowed in the demonstration version.")
			File_Save = False
			Exit Function
		End If
		On Error GoTo err_File_Save
		'SAVE FILE.
		''''f = FreeFile
		''''Open fn_Save For Output As #f
		''''Call File_Save_Latest_v1_40(f)
		frmMain.Cursor = System.Windows.Forms.Cursors.WaitCursor
		SavedOkay = File_Save_Latest_v1_60(fn_Save)
		If (SavedOkay = False) Then
			GoTo exit_err_File_Save
		End If
		''''Close #f
		'CLEAR DIRTY FLAG.
		Project_Is_Dirty = False
		Call DirtyStatus_Set_Current()
		File_Save = True
		'UPDATE DISPLAY.
		Filename = fn_Save
		frmMain.Text = AppName_For_Display_Short & "  -  " & Trim(Filename)
		Call frmMain_Refresh()
		'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
		Call OldFileList_Promote(Filename, 1, frmMain._mnuFileItem_199, frmMain.mnuFileItem(191), frmMain.mnuFileItem(192), frmMain.mnuFileItem(193), frmMain.mnuFileItem(194))
		GoTo Exit_Function
Exit_Function: 
		frmMain.Cursor = System.Windows.Forms.Cursors.Default
		Exit Function
exit_err_File_Save: 
		File_Save = False
		GoTo Exit_Function
err_File_Save: 
		Call Show_Trapped_Error("File_Save")
		On Error Resume Next
		FileClose(f)
		Resume exit_err_File_Save
	End Function
	'RETURNS:
	'- true = save went okay.
	'- false = save failed.
	Function File_SaveAs(ByRef fn_force As String) As Boolean
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim f As Short
		Dim fn_saveas As String
		Dim RetVal As Short
		If (IsThisADemo() = True) Then
			Call Demo_ShowError("Saving is not allowed in the demonstration version.")
			File_SaveAs = False
			Exit Function
		End If
		If (fn_force <> "") Then
			fn_saveas = fn_force
		Else
			Do While (1 = 1)
				'INPUT NEW FILENAME.
				On Error GoTo err_File_SaveAs
				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.CommonDialog1.DialogTitle = "Save " & AppName_For_Display_Short & " File"

				frmMain.SaveFileDialog1.Title = "Save " & AppName_For_Display_Short & " File"


				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.CommonDialog1.Filter = "All Files (*.*)|*.*|" & AppName_For_Display_Short & " Files (*.dat)|*.dat"

				frmMain.SaveFileDialog1.Filter = "All Files (*.*)|*.*|" & AppName_For_Display_Short & " Files (*.dat)|*.dat"


				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.CommonDialog1.FilterIndex = 2

				frmMain.SaveFileDialog1.FilterIndex = 2


				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.CommonDialog1.CancelError = True

				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNOverwritePrompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.CommonDialog1.flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist

				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.ShowSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.CommonDialog1.ShowSave()

				frmMain.SaveFileDialog1.ShowDialog()

				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fn_saveas = Trim(frmMain.SaveFileDialog1.FileName)
				If (LAUNCHFILEVIA_IS_DEBUG_MODE_ON) Then
					MsgBox("fn_saveas = `" & fn_saveas & "`")
				End If
				
				If (fn_saveas = "") Then
					'DO NOTHING.
					Exit Function
				End If
				'If (Not File_IsExists(fn_saveas)) Then
				'  Exit Do
				'End If
				'RetVal = MsgBox("File " & fn_saveas & _
				''    " already exists.  Do you want to replace it?", _
				''    vbQuestion + vbYesNo, _
				''    AppName_For_Display_Short & " : Overwrite File ?")
				'If (RetVal = vbYes) Then Exit Do
				'NOTE: "REPLACE?" CHECK HANDLED IN COMMON
				'DIALOG CONTROL NOW.
				Exit Do
			Loop 
		End If
		'OPEN THIS FILE.
		File_SaveAs = File_Save(fn_saveas)
		Exit Function
exit_err_File_SaveAs: 
		File_SaveAs = False
		Exit Function
err_File_SaveAs:
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = 75) Then
			'CANCEL BUTTON WAS SELECTED.
			Resume exit_err_File_SaveAs
		End If
		Call Show_Trapped_Error("File_SaveAs")
		Resume exit_err_File_SaveAs
	End Function
	
	
	Function Project_IsDirtyFlagThrown() As Boolean
		If (Project_Is_Dirty) Then
			Project_IsDirtyFlagThrown = True
		Else
			Project_IsDirtyFlagThrown = False
		End If
	End Function
	
	
	'RETURNS:
	'- true = it's okay to unload this file now.
	'- false = cancel the unload.
	Function file_query_unload() As Short
		Dim RetVal As Short
		Dim msg As String
		If (Not Project_IsDirtyFlagThrown()) Then
			file_query_unload = True
			Exit Function
		End If
		msg = "Do you want to save the changes you made to "
		If (Filename = "") Then
			msg = msg & "this new project"
		Else
			msg = msg & "your project of filename " & Filename
		End If
		msg = msg & " ?"
		RetVal = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNoCancel, AppName_For_Display_Short & " : Save Changes ?")
		Select Case RetVal
			Case MsgBoxResult.Yes
				If (File_SaveAs(Filename) = True) Then
					'SAVE WENT OK; IT'S NOW OKAY TO UNLOAD THIS FILE.
					file_query_unload = True
				Else
					'SAVE FAILED; DON'T UNLOAD THIS FILE.
					file_query_unload = False
				End If
				Exit Function
			Case MsgBoxResult.No
				file_query_unload = True
				Exit Function
			Case MsgBoxResult.Cancel
				file_query_unload = False
				Exit Function
		End Select
	End Function
	
	
	Sub ProjectFile_Read(ByRef f As Short, ByRef RetVal As Object, Optional ByRef optDummy1 As Object = Nothing)
		Dim outputstr As String
		Dim outlin As String
		Dim sub_name As String
		Dim input1 As String
		Dim input2 As String
		Input(f, input1)
		Input(f, input2)
		sub_name = "ProjectFile_Read"
		'UPGRADE_ISSUE: Constant vbDataObject was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(RetVal)
			Case VariantType.Boolean
				'UPGRADE_WARNING: Couldn't resolve default property of object RetVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RetVal = Val(input1)
			Case VariantType.Byte, VariantType.Short, VariantType.Integer, VariantType.Decimal
				'UPGRADE_WARNING: Couldn't resolve default property of object RetVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RetVal = Val(input1)
			Case VariantType.Single, VariantType.Double
				'UPGRADE_WARNING: Couldn't resolve default property of object RetVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RetVal = Val(input1)
			Case VariantType.String, VariantType.Date
				'UPGRADE_WARNING: Couldn't resolve default property of object RetVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RetVal = input1
			Case VariantType.Object
				MsgBox(sub_name & " vbObject not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Error
				MsgBox(sub_name & " vbError not implemented")
				GoTo EXIT_FALSE_VALUE
				'		Case vbDataObject
				'			MsgBox(sub_name & " vbDataObject not implemented")
	'			GoTo EXIT_FALSE_VALUE
			Case VariantType.Object
				MsgBox(sub_name & " vbVariant not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Array
				MsgBox(sub_name & " vbArray not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Empty
				MsgBox(sub_name & " vbEmpty not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Null
				MsgBox(sub_name & " vbNull not implemented")
				GoTo EXIT_FALSE_VALUE
		End Select
		GoTo EXIT_OK
EXIT_FALSE_VALUE: 
		PrintLine(f, "   - - - ERROR IN " & sub_name & "() - - -")
		Exit Sub
EXIT_OK: 
		Exit Sub
	End Sub
	Sub ProjectFile_Write(ByRef f As Short, ByRef v As Object, ByRef s As String)
		Dim outputstr As String
		Dim outlin As String
		Dim sub_name As String
		sub_name = "ProjectFile_Write"
		'UPGRADE_ISSUE: Constant vbDataObject was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(v)
			Case VariantType.Boolean
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = IIf(v, "1", "0")
			Case VariantType.Byte, VariantType.Short, VariantType.Integer, VariantType.Decimal
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = Trim(CStr(v))
			Case VariantType.Single, VariantType.Double
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = Trim(CStr(v))
			Case VariantType.String, VariantType.Date
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = CStr(v)
			Case VariantType.Object
				MsgBox(sub_name & " vbObject not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Error
				MsgBox(sub_name & " vbError not implemented")
				GoTo EXIT_FALSE_VALUE
				'	Case vbDataObject
				'			MsgBox(sub_name & " vbDataObject not implemented")
'				GoTo EXIT_FALSE_VALUE
			Case VariantType.Object
				MsgBox(sub_name & " vbVariant not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Array
				MsgBox(sub_name & " vbArray not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Empty
				MsgBox(sub_name & " vbEmpty not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Null
				MsgBox(sub_name & " vbNull not implemented")
				GoTo EXIT_FALSE_VALUE
		End Select
		outlin = Chr(34) & Trim(outputstr) & Chr(34) & "," & Chr(34) & s & Chr(34)
		'outlin = Trim$(outputstr$)
		'If (Len(outlin) > 27) Then
		'  outlin = outlin & "    "
		'Else
		'  Do While (1 = 1)
		'    If (Len(outlin) >= 27) Then Exit Do
		'    outlin = outlin & " "
		'  Loop
		'End If
		'outlin = outlin & s
		PrintLine(f, outlin)
		GoTo EXIT_OK
EXIT_FALSE_VALUE: 
		PrintLine(f, "   - - - ERROR IN " & sub_name & "() - - -")
		Exit Sub
EXIT_OK: 
		Exit Sub
	End Sub
End Module