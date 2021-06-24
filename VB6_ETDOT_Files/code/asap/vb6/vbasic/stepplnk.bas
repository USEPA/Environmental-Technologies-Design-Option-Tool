Attribute VB_Name = "SteppLnkMod"
Option Explicit

'
' THE ONLY UNCOMMENTED CODE IN THIS MODULE
' IS THE FOLLOWING GLOBAL VARIABLE:
'
Global StEPPImportSuccess As Integer



'
'Type StEPPLink_Property
'  Chemical As String
'  Cas As String
'  propname As String
'  units As String
'  val As Double       'If avail=False, val is 0.0E0
'  avail As Integer
'End Type
'
''---- Input variables to StEPP Link
'Global StEPPImportSuccess As Integer
'Global frmStEPPLink_Temperature As Double   'in DegC
'Global frmStEPPLink_Pressure As Double      'in Pa
'Global frmStEPPLink_ClientName As String    'name of client program, e.g. ADSIM
'Global StEPPLink_RequiredProps() As String  'list of strictly required properties
'Global StEPPLink_DontForget As String       'reminder for "Import Success" dialog
'Global StEPPLink_CurrentChemicalNames() As String
'
''---- Output variables from StEPP Link
'Global frmStEPPLink_Success As Integer
'Global StEPPLink_AllProps() As StEPPLink_Property
'Global StEPPLink_FilteredProps() As StEPPLink_Property
'Global StEPPLink_ImportFailed_Name() As String
'Global StEPPLink_ImportFailed_Reason() As String
'Global StEPPLink_ImportSucceeded_Name() As String
'
''---- Internal to StEPP Link
'Global fn_done_waitfile As String           'If this file exists, the StEPP link is complete
'Global fn_properties As String              'If the link was successful, this file contains the imported properties
'Global frmStEPPLink_SawRecreatedWaitfile As Integer   'Internal to StEPP Link
'
''Create a temporary file in the path {use_path}.
''Returns the filename {fn_temp}.
''Note: Does not return the path of the temporary file in {fn_temp}!
'Sub GetTempFilename(use_path As String, fn_temp As String)
'Dim temp As String
'Dim trycount As Integer
'Dim i As Integer
'Dim c As String
'Dim nowtime As String
'
'Dim save_path As String
'Dim f As Integer
'
'  save_path = CurDir$
'  ChDir use_path
'  ChDrive use_path
'
'  nowtime = Time$
'  temp = Left$(Time$, 2) + Mid$(Time$, 4, 2) + Right$(Time$, 2) + ".___"
'  trycount = 0
'  i = 1
'  Do While (1 = 1)
'    If (Dir(temp) = "") Then Exit Do
'    trycount = trycount + 1
'    'if (trycount > 40) then
'    i = i + 1
'    If (i >= 7) Then
'      i = 1
'    End If
'    c = Mid$(temp, i, 1)
'    If ((c >= "0") And (c <= "8")) Then
'      Mid$(temp, i, 1) = Chr$(Asc(c) + 1)
'    ElseIf ((c >= "A") And (c <= "Y")) Then
'      Mid$(temp, i, 1) = Chr$(Asc(c) + 1)
'    ElseIf (c = "9") Then
'      Mid$(temp, i, 1) = "A"
'    ElseIf (c = "Z") Then
'      Mid$(temp, i, 1) = "0"
'    End If
'  Loop
'
'  fn_temp = temp
'
'  f = FreeFile
'  Open fn_temp For Output As #f
'  Close #f
'  ChDir save_path
'  ChDrive save_path
'
'End Sub
'
'Sub StEPPLink_DisplayImportSucceeded()
'Dim num_import As Integer
'Dim num_failed As Integer
'Dim temp As String
'Dim i As Integer
'
'  num_import = UBound(StEPPLink_ImportSucceeded_Name)
'  num_failed = UBound(StEPPLink_ImportFailed_Name)
'
'  If (num_import <> 0) Then
'    temp = "Successfully imported " & Trim$(Str$(num_import)) & " component"
'    If (num_import <> 1) Then temp = temp & "s"
'    temp = temp & " from StEPP:"
'    For i = 1 To num_import
'      temp = temp & NL & "  " & Trim$(StEPPLink_ImportSucceeded_Name(i))
'    Next i
'    temp = temp & NL & "at pressure " & Trim$(Str$(frmStEPPLink_Pressure))
'    temp = temp & " Pa and temperature " & Trim$(Str$(frmStEPPLink_Temperature))
'    temp = temp & " Celcius."
'    temp = temp & NL
'    temp = temp & NL & StEPPLink_DontForget
'  Else
'    temp = "Unable to import the requested component(s)."
'  End If
'  If (num_failed <> 0) Then
'    temp = temp & NL
'    temp = temp & NL & "Failed to import the following " & Trim$(Str$(num_failed)) & " component"
'    If (num_failed <> 1) Then temp = temp & "s"
'    temp = temp & " from StEPP:"
'    For i = 1 To num_failed
'      temp = temp & NL & "  " & Trim$(StEPPLink_ImportFailed_Name(i))
'      temp = temp & " (unavailable properties: " & Trim$(StEPPLink_ImportFailed_Reason(i)) & ")"
'    Next i
'  End If
'  MsgBox temp, MB_ICONINFORMATION, Application_Name
'
'End Sub
'
'Sub StEPPLink_FilterUnimportable()
'Dim i As Integer
'Dim j As Integer
'Dim n As Integer
'Dim ub As Integer
'Dim num_failedimport As Integer
'Dim num_import As Integer
'Dim now_chemical As String
'Dim this_failed As Integer
'Dim this_failed_reason As String
'Dim importable() As Integer
'Dim s As String
'
'  '---- Misc inits
'  num_failedimport = 0
'  num_import = 0
'  now_chemical = ""
'
'  '---- Initialize success/failure arrays
'  ReDim StEPPLink_ImportSucceeded_Name(0 To 0)
'  ReDim StEPPLink_ImportFailed_Name(0 To 0)
'  ReDim StEPPLink_ImportFailed_Reason(0 To 0)
'
'  '---- Create arrays of successful and failed chemicals
'  ub = UBound(StEPPLink_AllProps)
'  For i = 1 To ub + 1
'    'If (i = ub + 1) Then
'    '  now_chemical = ""
'    'End If
'    If (i <> ub + 1) Then
'      s = StEPPLink_AllProps(i).Chemical
'    End If
'    If ((now_chemical <> s) Or (i = ub + 1)) Then
'      If (now_chemical <> "") Then
'        If (this_failed) Then
'          num_failedimport = num_failedimport + 1
'          If (UBound(StEPPLink_ImportFailed_Name) = 0) Then
'            ReDim StEPPLink_ImportFailed_Name(1 To 1)
'            ReDim StEPPLink_ImportFailed_Reason(1 To 1)
'          Else
'            ReDim Preserve StEPPLink_ImportFailed_Name(1 To num_failedimport)
'            ReDim Preserve StEPPLink_ImportFailed_Reason(1 To num_failedimport)
'          End If
'          StEPPLink_ImportFailed_Name(num_failedimport) = now_chemical
'          StEPPLink_ImportFailed_Reason(num_failedimport) = this_failed_reason
'        Else
'          num_import = num_import + 1
'          If (UBound(StEPPLink_ImportSucceeded_Name) = 0) Then
'            ReDim StEPPLink_ImportSucceeded_Name(1 To 1)
'          Else
'            ReDim Preserve StEPPLink_ImportSucceeded_Name(1 To num_import)
'          End If
'          StEPPLink_ImportSucceeded_Name(num_import) = now_chemical
'        End If
'      End If
'      If (i <> ub + 1) Then
'        now_chemical = StEPPLink_AllProps(i).Chemical
'        this_failed = False
'        this_failed_reason = ""
'        '-- Check to see if a chemical by this name already exists
'        For j = 1 To UBound(StEPPLink_CurrentChemicalNames)
'          If (Trim$(UCase$(now_chemical)) = Trim$(UCase$(StEPPLink_CurrentChemicalNames(j)))) Then
'            this_failed = True
'            this_failed_reason = "Component name already exists!"
'            Exit For
'          End If
'        Next j
'      End If
'    End If
'    If (i >= ub + 1) Then Exit For
'
'    If (Not StEPPLink_AllProps(i).avail) Then
'      '-- Check if this is a required property
'      For j = 1 To UBound(StEPPLink_RequiredProps)
'        If (UCase$(StEPPLink_RequiredProps(j)) = UCase$(StEPPLink_AllProps(i).propname)) Then
'          this_failed = True
'          If (Len(this_failed_reason) <> 0) Then this_failed_reason = this_failed_reason & ", "
'          this_failed_reason = this_failed_reason & StEPPLink_RequiredProps(j)
'          Exit For
'        End If
'      Next j
'    End If
'  Next i
'
'  '---- Create array of which properties have been filtered out (un-importable)
'  ReDim importable(1 To UBound(StEPPLink_AllProps))
'  ub = UBound(StEPPLink_AllProps)
'  For i = 1 To ub
'    importable(i) = False
'    For j = 1 To num_import
'      If (UCase$(StEPPLink_AllProps(i).Chemical) = UCase$(StEPPLink_ImportSucceeded_Name(j))) Then
'        importable(i) = True
'        Exit For
'      End If
'    Next j
'  Next i
'
'  '---- Output importable properties to StEPPLink_FilteredProps()
'  n = 0
'  For i = 1 To ub
'    If (importable(i)) Then
'      n = n + 1
'      ReDim Preserve StEPPLink_FilteredProps(1 To n)
'      StEPPLink_FilteredProps(n) = StEPPLink_AllProps(i)
'    End If
'  Next i
'
'End Sub
'
''Note: Returns -1 if the property cannot be found!
'Function StEPPLink_FindProp(chem As String, propname As String) As Integer
'Dim i As Integer
'Dim ub As Integer
'Dim s As String
'
'  ub = UBound(StEPPLink_FilteredProps)
'  For i = 1 To ub
'    s = StEPPLink_FilteredProps(i).Chemical
'    If (UCase$(chem) = UCase$(s)) Then
'      s = StEPPLink_FilteredProps(i).propname
'      If (UCase$(propname) = UCase$(s)) Then
'        StEPPLink_FindProp = i
'        Exit Function
'      End If
'    End If
'  Next i
'
'  StEPPLink_FindProp = -1
'
'End Function
'
'Sub StEPPLink_ImportPropertyFile(fn As String)
'Dim f As Integer
'Dim s1 As String
'Dim s2 As String
'Dim s3 As String
'Dim n As Integer
'Dim now_chemical As String
'Dim now_cas As String
'
'  n = 0
'  f = FreeFile
'  Open fn For Input As #f
'  Do While (1 = 1)
'    If (EOF(f)) Then Exit Do
'    Input #f, s1, s2, s3
'    If (s1 = "END_OF_FILE") Then Exit Do
'    If (UCase$(s1) = "CHEMICAL") Then
'      now_chemical = s2
'      now_cas = s3
'    Else
'      n = n + 1
'      ReDim Preserve StEPPLink_AllProps(1 To n)
'      StEPPLink_AllProps(n).Chemical = now_chemical
'      StEPPLink_AllProps(n).Cas = now_cas
'      StEPPLink_AllProps(n).propname = s1
'      StEPPLink_AllProps(n).units = s2
'      If (UCase$(s3) <> "UNAVAILABLE") Then
'        StEPPLink_AllProps(n).val = CDbl(s3)
'        StEPPLink_AllProps(n).avail = True
'      Else
'        StEPPLink_AllProps(n).val = 0#
'        StEPPLink_AllProps(n).avail = False
'      End If
'    End If
'  Loop
'
'  Close #f
'
'End Sub
'
