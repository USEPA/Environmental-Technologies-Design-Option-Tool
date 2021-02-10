<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmModelCPHSDMResults
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public optType(3) As AxThreed.AxSSOption
    Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents Command4 As System.Windows.Forms.Button
    Public WithEvents grpBreak As AxGraphLib.AxGraph
    Public WithEvents cboGrid As System.Windows.Forms.ComboBox

    Public WithEvents cmdTreat As AxThreed.AxSSCommand
    Public WithEvents _lblLegend_7 As System.Windows.Forms.Label
    Public WithEvents _lblData_11 As System.Windows.Forms.Label
    Public WithEvents _lblData_10 As System.Windows.Forms.Label
    Public WithEvents _lblData_9 As System.Windows.Forms.Label
    Public WithEvents _lblData_8 As System.Windows.Forms.Label
    Public WithEvents _lblData_7 As System.Windows.Forms.Label
    Public WithEvents _lblData_6 As System.Windows.Forms.Label
    Public WithEvents _lblData_5 As System.Windows.Forms.Label
    Public WithEvents _lblData_4 As System.Windows.Forms.Label
    Public WithEvents _lblData_3 As System.Windows.Forms.Label
    Public WithEvents _lblData_2 As System.Windows.Forms.Label
    Public WithEvents _lblData_1 As System.Windows.Forms.Label
    Public WithEvents _lblData_0 As System.Windows.Forms.Label
    Public WithEvents _lblLegend_6 As System.Windows.Forms.Label
    Public WithEvents _lblLegend_5 As System.Windows.Forms.Label
    Public WithEvents _lblLegend_4 As System.Windows.Forms.Label
    Public WithEvents _lblLegend_3 As System.Windows.Forms.Label
    Public WithEvents _lblLegend_2 As System.Windows.Forms.Label
    Public WithEvents _lblLegend_1 As System.Windows.Forms.Label
    Public WithEvents _lblLegend_0 As System.Windows.Forms.Label
    Public WithEvents _lblPara_0 As System.Windows.Forms.Label
    Public WithEvents _lblPara_1 As System.Windows.Forms.Label
    Public WithEvents _lblPara_2 As System.Windows.Forms.Label
    Public WithEvents _lblPara_6 As System.Windows.Forms.Label
    Public WithEvents _lblParaValue_0 As System.Windows.Forms.Label
    Public WithEvents _lblParaValue_2 As System.Windows.Forms.Label
    Public WithEvents _lblParaValue_5 As System.Windows.Forms.Label
    Public WithEvents _lblParaValue_6 As System.Windows.Forms.Label
    Public WithEvents _lblData_15 As System.Windows.Forms.Label
    Public WithEvents _lblData_16 As System.Windows.Forms.Label
    Public WithEvents _lblData_17 As System.Windows.Forms.Label
    Public WithEvents _lblData_18 As System.Windows.Forms.Label
    Public WithEvents Frame3D1 As AxThreed.AxSSFrame
    Public WithEvents cmdFile As AxThreed.AxSSCommand
    Public WithEvents cmdExcel As AxThreed.AxSSCommand
    Public WithEvents CMDialog1 As AxMSComDlg.AxCommonDialog
    Public WithEvents cmdSave As AxThreed.AxSSCommand
    Public WithEvents cmdSelect As AxThreed.AxSSCommand
    Public WithEvents cmdPrint As AxThreed.AxSSCommand
    Public WithEvents cmdExit As AxThreed.AxSSCommand
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblData As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblLegend As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblPara As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblParaValue As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    '   Public WithEvents optType As Threed.SSOptionArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmModelCPHSDMResults))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.cboGrid = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblData = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblData_9 = New System.Windows.Forms.Label()
        Me._lblData_6 = New System.Windows.Forms.Label()
        Me._lblData_3 = New System.Windows.Forms.Label()
        Me._lblData_0 = New System.Windows.Forms.Label()
        Me._lblData_10 = New System.Windows.Forms.Label()
        Me._lblData_7 = New System.Windows.Forms.Label()
        Me._lblData_4 = New System.Windows.Forms.Label()
        Me._lblData_1 = New System.Windows.Forms.Label()
        Me._lblData_11 = New System.Windows.Forms.Label()
        Me._lblData_8 = New System.Windows.Forms.Label()
        Me._lblData_5 = New System.Windows.Forms.Label()
        Me._lblData_2 = New System.Windows.Forms.Label()
        Me._lblData_18 = New System.Windows.Forms.Label()
        Me._lblData_17 = New System.Windows.Forms.Label()
        Me._lblData_16 = New System.Windows.Forms.Label()
        Me._lblData_15 = New System.Windows.Forms.Label()
        Me.lblLegend = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblLegend_0 = New System.Windows.Forms.Label()
        Me._lblLegend_4 = New System.Windows.Forms.Label()
        Me._lblLegend_5 = New System.Windows.Forms.Label()
        Me._lblLegend_6 = New System.Windows.Forms.Label()
        Me._lblLegend_1 = New System.Windows.Forms.Label()
        Me._lblLegend_2 = New System.Windows.Forms.Label()
        Me._lblLegend_3 = New System.Windows.Forms.Label()
        Me._lblLegend_7 = New System.Windows.Forms.Label()
        Me.lblPara = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblPara_0 = New System.Windows.Forms.Label()
        Me._lblPara_1 = New System.Windows.Forms.Label()
        Me._lblPara_2 = New System.Windows.Forms.Label()
        Me._lblPara_6 = New System.Windows.Forms.Label()
        Me.lblParaValue = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblParaValue_0 = New System.Windows.Forms.Label()
        Me._lblParaValue_5 = New System.Windows.Forms.Label()
        Me._lblParaValue_6 = New System.Windows.Forms.Label()
        Me._lblParaValue_2 = New System.Windows.Forms.Label()
        Me.grpBreak = New AxGraphLib.AxGraph()
        Me.CMDialog1 = New AxMSComDlg.AxCommonDialog()
        Me.Frame3D1 = New AxThreed.AxSSFrame()
        Me.cmdFile = New AxThreed.AxSSCommand()
        Me.cmdExcel = New AxThreed.AxSSCommand()
        Me.cmdSave = New AxThreed.AxSSCommand()
        Me.cmdSelect = New AxThreed.AxSSCommand()
        Me.cmdPrint = New AxThreed.AxSSCommand()
        Me.cmdExit = New AxThreed.AxSSCommand()
        Me.cmdTreat = New AxThreed.AxSSCommand()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me._optType_2 = New System.Windows.Forms.RadioButton()
        Me._optType_1 = New System.Windows.Forms.RadioButton()
        Me._optType_0 = New System.Windows.Forms.RadioButton()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLegend, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPara, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblParaValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpBreak, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CMDialog1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame3D1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdFile, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdExcel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdSave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdSelect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdExit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdTreat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(531, 433)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(110, 22)
        Me.Command4.TabIndex = 47
        Me.Command4.Text = "Print Screen"
        Me.ToolTip1.SetToolTip(Me.Command4, "Click here to print current screen to selected printer")
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(658, 386)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 48
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'cboGrid
        '
        Me.cboGrid.BackColor = System.Drawing.SystemColors.Window
        Me.cboGrid.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboGrid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGrid.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboGrid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboGrid.Location = New System.Drawing.Point(531, 176)
        Me.cboGrid.Name = "cboGrid"
        Me.cboGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboGrid.Size = New System.Drawing.Size(100, 24)
        Me.cboGrid.TabIndex = 41
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(528, 156)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(105, 17)
        Me.Label1.TabIndex = 44
        Me.Label1.Text = "Grid Style:"
        '
        '_lblData_9
        '
        Me._lblData_9.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_9, CType(9, Short))
        Me._lblData_9.Location = New System.Drawing.Point(455, 111)
        Me._lblData_9.Name = "_lblData_9"
        Me._lblData_9.Size = New System.Drawing.Size(66, 18)
        Me._lblData_9.TabIndex = 74
        Me._lblData_9.Text = " 999999 "
        '
        '_lblData_6
        '
        Me._lblData_6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_6, CType(6, Short))
        Me._lblData_6.Location = New System.Drawing.Point(370, 111)
        Me._lblData_6.Name = "_lblData_6"
        Me._lblData_6.Size = New System.Drawing.Size(86, 18)
        Me._lblData_6.TabIndex = 73
        Me._lblData_6.Text = "  9999999   "
        '
        '_lblData_3
        '
        Me._lblData_3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_3, CType(3, Short))
        Me._lblData_3.Location = New System.Drawing.Point(309, 111)
        Me._lblData_3.Name = "_lblData_3"
        Me._lblData_3.Size = New System.Drawing.Size(62, 18)
        Me._lblData_3.TabIndex = 72
        Me._lblData_3.Text = "   999    "
        '
        '_lblData_0
        '
        Me._lblData_0.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_0, CType(0, Short))
        Me._lblData_0.Location = New System.Drawing.Point(226, 111)
        Me._lblData_0.Name = "_lblData_0"
        Me._lblData_0.Size = New System.Drawing.Size(90, 18)
        Me._lblData_0.TabIndex = 71
        Me._lblData_0.Text = "  99999999  "
        '
        '_lblData_10
        '
        Me._lblData_10.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_10, CType(10, Short))
        Me._lblData_10.Location = New System.Drawing.Point(455, 130)
        Me._lblData_10.Name = "_lblData_10"
        Me._lblData_10.Size = New System.Drawing.Size(66, 18)
        Me._lblData_10.TabIndex = 78
        Me._lblData_10.Text = " 999999 "
        '
        '_lblData_7
        '
        Me._lblData_7.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_7, CType(7, Short))
        Me._lblData_7.Location = New System.Drawing.Point(370, 130)
        Me._lblData_7.Name = "_lblData_7"
        Me._lblData_7.Size = New System.Drawing.Size(86, 18)
        Me._lblData_7.TabIndex = 77
        Me._lblData_7.Text = "  9999999   "
        '
        '_lblData_4
        '
        Me._lblData_4.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_4, CType(4, Short))
        Me._lblData_4.Location = New System.Drawing.Point(309, 130)
        Me._lblData_4.Name = "_lblData_4"
        Me._lblData_4.Size = New System.Drawing.Size(62, 18)
        Me._lblData_4.TabIndex = 76
        Me._lblData_4.Text = "   999    "
        '
        '_lblData_1
        '
        Me._lblData_1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_1, CType(1, Short))
        Me._lblData_1.Location = New System.Drawing.Point(226, 130)
        Me._lblData_1.Name = "_lblData_1"
        Me._lblData_1.Size = New System.Drawing.Size(90, 18)
        Me._lblData_1.TabIndex = 75
        Me._lblData_1.Text = "  99999999  "
        '
        '_lblData_11
        '
        Me._lblData_11.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_11, CType(11, Short))
        Me._lblData_11.Location = New System.Drawing.Point(455, 148)
        Me._lblData_11.Name = "_lblData_11"
        Me._lblData_11.Size = New System.Drawing.Size(66, 18)
        Me._lblData_11.TabIndex = 82
        Me._lblData_11.Text = " 999999 "
        '
        '_lblData_8
        '
        Me._lblData_8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_8, CType(8, Short))
        Me._lblData_8.Location = New System.Drawing.Point(370, 148)
        Me._lblData_8.Name = "_lblData_8"
        Me._lblData_8.Size = New System.Drawing.Size(86, 18)
        Me._lblData_8.TabIndex = 81
        Me._lblData_8.Text = "  9999999   "
        '
        '_lblData_5
        '
        Me._lblData_5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_5, CType(5, Short))
        Me._lblData_5.Location = New System.Drawing.Point(309, 148)
        Me._lblData_5.Name = "_lblData_5"
        Me._lblData_5.Size = New System.Drawing.Size(62, 18)
        Me._lblData_5.TabIndex = 80
        Me._lblData_5.Text = "   999    "
        '
        '_lblData_2
        '
        Me._lblData_2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_2, CType(2, Short))
        Me._lblData_2.Location = New System.Drawing.Point(226, 148)
        Me._lblData_2.Name = "_lblData_2"
        Me._lblData_2.Size = New System.Drawing.Size(90, 18)
        Me._lblData_2.TabIndex = 79
        Me._lblData_2.Text = "  99999999  "
        '
        '_lblData_18
        '
        Me._lblData_18.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_18, CType(18, Short))
        Me._lblData_18.Location = New System.Drawing.Point(455, 165)
        Me._lblData_18.Name = "_lblData_18"
        Me._lblData_18.Size = New System.Drawing.Size(66, 18)
        Me._lblData_18.TabIndex = 86
        Me._lblData_18.Text = " 999999 "
        '
        '_lblData_17
        '
        Me._lblData_17.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_17, CType(17, Short))
        Me._lblData_17.Location = New System.Drawing.Point(370, 165)
        Me._lblData_17.Name = "_lblData_17"
        Me._lblData_17.Size = New System.Drawing.Size(86, 18)
        Me._lblData_17.TabIndex = 85
        Me._lblData_17.Text = "  9999999   "
        '
        '_lblData_16
        '
        Me._lblData_16.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_16, CType(16, Short))
        Me._lblData_16.Location = New System.Drawing.Point(309, 165)
        Me._lblData_16.Name = "_lblData_16"
        Me._lblData_16.Size = New System.Drawing.Size(62, 18)
        Me._lblData_16.TabIndex = 84
        Me._lblData_16.Text = "   999    "
        '
        '_lblData_15
        '
        Me._lblData_15.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me._lblData_15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData.SetIndex(Me._lblData_15, CType(15, Short))
        Me._lblData_15.Location = New System.Drawing.Point(226, 165)
        Me._lblData_15.Name = "_lblData_15"
        Me._lblData_15.Size = New System.Drawing.Size(90, 18)
        Me._lblData_15.TabIndex = 83
        Me._lblData_15.Text = "  99999999  "
        '
        '_lblLegend_0
        '
        Me._lblLegend_0.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_0, CType(0, Short))
        Me._lblLegend_0.Location = New System.Drawing.Point(37, 93)
        Me._lblLegend_0.Name = "_lblLegend_0"
        Me._lblLegend_0.Size = New System.Drawing.Size(190, 18)
        Me._lblLegend_0.TabIndex = 62
        Me._lblLegend_0.Text = "                                             "
        '
        '_lblLegend_4
        '
        Me._lblLegend_4.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_4, CType(4, Short))
        Me._lblLegend_4.Location = New System.Drawing.Point(37, 111)
        Me._lblLegend_4.Name = "_lblLegend_4"
        Me._lblLegend_4.Size = New System.Drawing.Size(189, 18)
        Me._lblLegend_4.TabIndex = 63
        Me._lblLegend_4.Text = "5% of Influent Concen.         "
        '
        '_lblLegend_5
        '
        Me._lblLegend_5.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_5, CType(5, Short))
        Me._lblLegend_5.Location = New System.Drawing.Point(37, 129)
        Me._lblLegend_5.Name = "_lblLegend_5"
        Me._lblLegend_5.Size = New System.Drawing.Size(189, 18)
        Me._lblLegend_5.TabIndex = 64
        Me._lblLegend_5.Text = "50% of Influent Concen.       "
        '
        '_lblLegend_6
        '
        Me._lblLegend_6.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_6, CType(6, Short))
        Me._lblLegend_6.Location = New System.Drawing.Point(37, 147)
        Me._lblLegend_6.Name = "_lblLegend_6"
        Me._lblLegend_6.Size = New System.Drawing.Size(189, 18)
        Me._lblLegend_6.TabIndex = 65
        Me._lblLegend_6.Text = "95% of Influent Concen.       "
        '
        '_lblLegend_1
        '
        Me._lblLegend_1.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_1, CType(1, Short))
        Me._lblLegend_1.Location = New System.Drawing.Point(227, 93)
        Me._lblLegend_1.Name = "_lblLegend_1"
        Me._lblLegend_1.Size = New System.Drawing.Size(88, 18)
        Me._lblLegend_1.TabIndex = 67
        Me._lblLegend_1.Text = "Time(days)  "
        '
        '_lblLegend_2
        '
        Me._lblLegend_2.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_2, CType(2, Short))
        Me._lblLegend_2.Location = New System.Drawing.Point(310, 93)
        Me._lblLegend_2.Name = "_lblLegend_2"
        Me._lblLegend_2.Size = New System.Drawing.Size(61, 18)
        Me._lblLegend_2.TabIndex = 68
        Me._lblLegend_2.Text = "   BVT   "
        '
        '_lblLegend_3
        '
        Me._lblLegend_3.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_3, CType(3, Short))
        Me._lblLegend_3.Location = New System.Drawing.Point(371, 93)
        Me._lblLegend_3.Name = "_lblLegend_3"
        Me._lblLegend_3.Size = New System.Drawing.Size(85, 18)
        Me._lblLegend_3.TabIndex = 69
        Me._lblLegend_3.Text = "Tr. Capacity"
        '
        '_lblLegend_7
        '
        Me._lblLegend_7.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me._lblLegend_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLegend.SetIndex(Me._lblLegend_7, CType(7, Short))
        Me._lblLegend_7.Location = New System.Drawing.Point(456, 93)
        Me._lblLegend_7.Name = "_lblLegend_7"
        Me._lblLegend_7.Size = New System.Drawing.Size(65, 18)
        Me._lblLegend_7.TabIndex = 70
        Me._lblLegend_7.Text = "C (mg/L)"
        '
        '_lblPara_0
        '
        Me._lblPara_0.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.lblPara.SetIndex(Me._lblPara_0, CType(0, Short))
        Me._lblPara_0.Location = New System.Drawing.Point(37, 35)
        Me._lblPara_0.Name = "_lblPara_0"
        Me._lblPara_0.Size = New System.Drawing.Size(195, 16)
        Me._lblPara_0.TabIndex = 54
        Me._lblPara_0.Text = "Minimum Stanton Number"
        '
        '_lblPara_1
        '
        Me._lblPara_1.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.lblPara.SetIndex(Me._lblPara_1, CType(1, Short))
        Me._lblPara_1.Location = New System.Drawing.Point(37, 51)
        Me._lblPara_1.Name = "_lblPara_1"
        Me._lblPara_1.Size = New System.Drawing.Size(195, 16)
        Me._lblPara_1.TabIndex = 56
        Me._lblPara_1.Text = "Minimum EBCT (min)"
        '
        '_lblPara_2
        '
        Me._lblPara_2.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.lblPara.SetIndex(Me._lblPara_2, CType(2, Short))
        Me._lblPara_2.Location = New System.Drawing.Point(37, 67)
        Me._lblPara_2.Name = "_lblPara_2"
        Me._lblPara_2.Size = New System.Drawing.Size(195, 16)
        Me._lblPara_2.TabIndex = 58
        Me._lblPara_2.Text = "Minimum Column Length (cm)"
        '
        '_lblPara_6
        '
        Me._lblPara_6.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.lblPara.SetIndex(Me._lblPara_6, CType(6, Short))
        Me._lblPara_6.Location = New System.Drawing.Point(307, 67)
        Me._lblPara_6.Name = "_lblPara_6"
        Me._lblPara_6.Size = New System.Drawing.Size(116, 16)
        Me._lblPara_6.TabIndex = 60
        Me._lblPara_6.Text = "MTZ Length (cm)"
        '
        '_lblParaValue_0
        '
        Me.lblParaValue.SetIndex(Me._lblParaValue_0, CType(0, Short))
        Me._lblParaValue_0.Location = New System.Drawing.Point(238, 35)
        Me._lblParaValue_0.Name = "_lblParaValue_0"
        Me._lblParaValue_0.Size = New System.Drawing.Size(51, 16)
        Me._lblParaValue_0.TabIndex = 55
        Me._lblParaValue_0.Text = "Label4"
        '
        '_lblParaValue_5
        '
        Me.lblParaValue.SetIndex(Me._lblParaValue_5, CType(5, Short))
        Me._lblParaValue_5.Location = New System.Drawing.Point(238, 51)
        Me._lblParaValue_5.Name = "_lblParaValue_5"
        Me._lblParaValue_5.Size = New System.Drawing.Size(51, 16)
        Me._lblParaValue_5.TabIndex = 57
        Me._lblParaValue_5.Text = "Label2"
        '
        '_lblParaValue_6
        '
        Me.lblParaValue.SetIndex(Me._lblParaValue_6, CType(6, Short))
        Me._lblParaValue_6.Location = New System.Drawing.Point(238, 67)
        Me._lblParaValue_6.Name = "_lblParaValue_6"
        Me._lblParaValue_6.Size = New System.Drawing.Size(51, 16)
        Me._lblParaValue_6.TabIndex = 59
        Me._lblParaValue_6.Text = "Label6"
        '
        '_lblParaValue_2
        '
        Me.lblParaValue.SetIndex(Me._lblParaValue_2, CType(2, Short))
        Me._lblParaValue_2.Location = New System.Drawing.Point(445, 67)
        Me._lblParaValue_2.Name = "_lblParaValue_2"
        Me._lblParaValue_2.Size = New System.Drawing.Size(51, 16)
        Me._lblParaValue_2.TabIndex = 61
        Me._lblParaValue_2.Text = "Label8"
        '
        'grpBreak
        '
        Me.grpBreak.Location = New System.Drawing.Point(24, 206)
        Me.grpBreak.Name = "grpBreak"
        Me.grpBreak.OcxState = CType(resources.GetObject("grpBreak.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grpBreak.Size = New System.Drawing.Size(489, 249)
        Me.grpBreak.TabIndex = 46
        '
        'CMDialog1
        '
        Me.CMDialog1.Enabled = True
        Me.CMDialog1.Location = New System.Drawing.Point(582, 80)
        Me.CMDialog1.Name = "CMDialog1"
        Me.CMDialog1.OcxState = CType(resources.GetObject("CMDialog1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CMDialog1.Size = New System.Drawing.Size(40, 40)
        Me.CMDialog1.TabIndex = 49
        '
        'Frame3D1
        '
        Me.Frame3D1.Location = New System.Drawing.Point(24, 12)
        Me.Frame3D1.Name = "Frame3D1"
        Me.Frame3D1.OcxState = CType(resources.GetObject("Frame3D1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Frame3D1.Size = New System.Drawing.Size(501, 188)
        Me.Frame3D1.TabIndex = 3
        '
        'cmdFile
        '
        Me.cmdFile.Location = New System.Drawing.Point(531, 368)
        Me.cmdFile.Name = "cmdFile"
        Me.cmdFile.OcxState = CType(resources.GetObject("cmdFile.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdFile.Size = New System.Drawing.Size(100, 44)
        Me.cmdFile.TabIndex = 38
        Me.cmdFile.TabStop = False
        '
        'cmdExcel
        '
        Me.cmdExcel.Location = New System.Drawing.Point(531, 206)
        Me.cmdExcel.Name = "cmdExcel"
        Me.cmdExcel.OcxState = CType(resources.GetObject("cmdExcel.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdExcel.Size = New System.Drawing.Size(100, 44)
        Me.cmdExcel.TabIndex = 39
        Me.cmdExcel.TabStop = False
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(531, 247)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.OcxState = CType(resources.GetObject("cmdSave.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdSave.Size = New System.Drawing.Size(100, 44)
        Me.cmdSave.TabIndex = 40
        Me.cmdSave.TabStop = False
        '
        'cmdSelect
        '
        Me.cmdSelect.Location = New System.Drawing.Point(531, 287)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.OcxState = CType(resources.GetObject("cmdSelect.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdSelect.Size = New System.Drawing.Size(100, 44)
        Me.cmdSelect.TabIndex = 42
        Me.cmdSelect.TabStop = False
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(531, 327)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.OcxState = CType(resources.GetObject("cmdPrint.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdPrint.Size = New System.Drawing.Size(100, 44)
        Me.cmdPrint.TabIndex = 43
        Me.cmdPrint.TabStop = False
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(555, 17)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.OcxState = CType(resources.GetObject("cmdExit.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdExit.Size = New System.Drawing.Size(100, 50)
        Me.cmdExit.TabIndex = 45
        Me.cmdExit.TabStop = False
        '
        'cmdTreat
        '
        Me.cmdTreat.Location = New System.Drawing.Point(37, 165)
        Me.cmdTreat.Name = "cmdTreat"
        Me.cmdTreat.OcxState = CType(resources.GetObject("cmdTreat.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdTreat.Size = New System.Drawing.Size(189, 18)
        Me.cmdTreat.TabIndex = 87
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me._optType_2)
        Me.GroupBox1.Controls.Add(Me._optType_1)
        Me.GroupBox1.Controls.Add(Me._optType_0)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 461)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(399, 59)
        Me.GroupBox1.TabIndex = 91
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "C/C0 As a Function of:"
        '
        '_optType_2
        '
        Me._optType_2.AutoSize = True
        Me._optType_2.Location = New System.Drawing.Point(177, 19)
        Me._optType_2.Name = "_optType_2"
        Me._optType_2.Size = New System.Drawing.Size(184, 20)
        Me._optType_2.TabIndex = 2
        Me._optType_2.TabStop = True
        Me._optType_2.Text = "Volume Treated by Mass"
        Me._optType_2.UseVisualStyleBackColor = True
        '
        '_optType_1
        '
        Me._optType_1.AutoSize = True
        Me._optType_1.Location = New System.Drawing.Point(92, 19)
        Me._optType_1.Name = "_optType_1"
        Me._optType_1.Size = New System.Drawing.Size(56, 20)
        Me._optType_1.TabIndex = 1
        Me._optType_1.Text = "BVT"
        Me._optType_1.UseVisualStyleBackColor = True
        '
        '_optType_0
        '
        Me._optType_0.AutoSize = True
        Me._optType_0.Checked = True
        Me._optType_0.Location = New System.Drawing.Point(8, 19)
        Me._optType_0.Name = "_optType_0"
        Me._optType_0.Size = New System.Drawing.Size(59, 20)
        Me._optType_0.TabIndex = 0
        Me._optType_0.TabStop = True
        Me._optType_0.Text = "Time"
        Me._optType_0.UseVisualStyleBackColor = True
        '
        'frmModelCPHSDMResults
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(661, 532)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdTreat)
        Me.Controls.Add(Me._lblData_18)
        Me.Controls.Add(Me._lblData_17)
        Me.Controls.Add(Me._lblData_16)
        Me.Controls.Add(Me._lblData_15)
        Me.Controls.Add(Me._lblData_11)
        Me.Controls.Add(Me._lblData_8)
        Me.Controls.Add(Me._lblData_5)
        Me.Controls.Add(Me._lblData_2)
        Me.Controls.Add(Me._lblData_10)
        Me.Controls.Add(Me._lblData_7)
        Me.Controls.Add(Me._lblData_4)
        Me.Controls.Add(Me._lblData_1)
        Me.Controls.Add(Me._lblData_9)
        Me.Controls.Add(Me._lblData_6)
        Me.Controls.Add(Me._lblData_3)
        Me.Controls.Add(Me._lblData_0)
        Me.Controls.Add(Me._lblLegend_7)
        Me.Controls.Add(Me._lblLegend_3)
        Me.Controls.Add(Me._lblLegend_2)
        Me.Controls.Add(Me._lblLegend_1)
        Me.Controls.Add(Me._lblLegend_6)
        Me.Controls.Add(Me._lblLegend_5)
        Me.Controls.Add(Me._lblLegend_4)
        Me.Controls.Add(Me._lblLegend_0)
        Me.Controls.Add(Me._lblParaValue_2)
        Me.Controls.Add(Me._lblPara_6)
        Me.Controls.Add(Me._lblParaValue_6)
        Me.Controls.Add(Me._lblPara_2)
        Me.Controls.Add(Me._lblParaValue_5)
        Me.Controls.Add(Me._lblPara_1)
        Me.Controls.Add(Me._lblParaValue_0)
        Me.Controls.Add(Me._lblPara_0)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.grpBreak)
        Me.Controls.Add(Me.cboGrid)
        Me.Controls.Add(Me.Frame3D1)
        Me.Controls.Add(Me.cmdFile)
        Me.Controls.Add(Me.cmdExcel)
        Me.Controls.Add(Me.CMDialog1)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdSelect)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(157, 58)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmModelCPHSDMResults"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Results for the Constant Pattern Model (CPHSDM)"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLegend, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPara, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblParaValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpBreak, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CMDialog1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame3D1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdFile, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdExcel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdSave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdSelect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdExit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdTreat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents _optType_2 As RadioButton
    Friend WithEvents _optType_1 As RadioButton
    Friend WithEvents _optType_0 As RadioButton

    'Friend WithEvents cmdTreatA As AxThreed.AxSSCommand
    '  Friend WithEvents _lblLegend_1A As Label
    '   Friend WithEvents _lblLegend_2A As Label
    '   Friend WithEvents _lblLegend_3A As Label
    '   Friend WithEvents _lblLegend_7A As Label
    ' Friend WithEvents _lblData_9A As Label
    '  Friend WithEvents _lblData_6A As Label
    '   Friend WithEvents _lblData_3 As Label
    '   Friend WithEvents _lblData_0A As Label
    ' Friend WithEvents _lblData_10A As Label
    '   Friend WithEvents _lblData_7A As Label
    '  Friend WithEvents _lblData_4A As Label
    ' Friend WithEvents _lblData_1A As Label
    ' Friend WithEvents _lblData_11A As Label
    ' Friend WithEvents _lblData_8A As Label
    '  Friend WithEvents _lblData_5A As Label
    '  Friend WithEvents _lblData_2A As Label
    '   Friend WithEvents _lblData_18A As Label
    '  Friend WithEvents _lblData_17A As Label
    ' Friend WithEvents _lblData_16A As Label
    '  Friend WithEvents _lblData_15A As Label
    '    Public WithEvents AxSSOption1 As AxThreed.AxSSOption
    '    Public WithEvents AxSSOption2 As AxThreed.AxSSOption
    '   Public WithEvents AxSSOption3 As AxThreed.AxSSOption
    '  Public WithEvents AxSSFrame1 As AxThreed.AxSSFrame

#End Region
End Class