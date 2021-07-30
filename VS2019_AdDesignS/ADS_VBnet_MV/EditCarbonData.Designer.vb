<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEditCarbonData
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
    '   Public WithEvents cmdSaveCancel As SSCommandArray
    Public optPhase(2) As AxThreed.AxSSOption
    Public cmdSaveCancel(3) As AxThreed.AxSSCommand
    Public WithEvents lblDesc As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    '   Public WithEvents optPhase As SSOptionArray
    Public WithEvents txtData As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblDesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblDesc_7 = New System.Windows.Forms.Label()
        Me._lblDesc_6 = New System.Windows.Forms.Label()
        Me._lblDesc_5 = New System.Windows.Forms.Label()
        Me._lblDesc_4 = New System.Windows.Forms.Label()
        Me._lblDesc_3 = New System.Windows.Forms.Label()
        Me._lblDesc_2 = New System.Windows.Forms.Label()
        Me._lblDesc_1 = New System.Windows.Forms.Label()
        Me._lblDesc_0 = New System.Windows.Forms.Label()
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblUnit_6 = New System.Windows.Forms.Label()
        Me._lblUnit_5 = New System.Windows.Forms.Label()
        Me._lblUnit_4 = New System.Windows.Forms.Label()
        Me._lblUnit_2 = New System.Windows.Forms.Label()
        Me._lblUnit_1 = New System.Windows.Forms.Label()
        Me._lblUnit_0 = New System.Windows.Forms.Label()
        Me.txtData = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me._txtData_7 = New System.Windows.Forms.TextBox()
        Me._txtData_6 = New System.Windows.Forms.TextBox()
        Me._txtData_5 = New System.Windows.Forms.TextBox()
        Me._txtData_4 = New System.Windows.Forms.TextBox()
        Me._txtData_3 = New System.Windows.Forms.TextBox()
        Me._txtData_2 = New System.Windows.Forms.TextBox()
        Me._txtData_1 = New System.Windows.Forms.TextBox()
        Me._txtData_0 = New System.Windows.Forms.TextBox()
        Me._optPhase_1 = New System.Windows.Forms.RadioButton()
        Me._optPhase_2 = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblUnitB = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me._cmdSaveCancel_0 = New System.Windows.Forms.Button()
        Me._cmdSaveCancel_1 = New System.Windows.Forms.Button()
        Me._cmdSaveCancel_2 = New System.Windows.Forms.Button()
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        '_lblDesc_7
        '
        Me._lblDesc_7.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_7, CType(7, Short))
        Me._lblDesc_7.Location = New System.Drawing.Point(-4, 30)
        Me._lblDesc_7.Name = "_lblDesc_7"
        Me._lblDesc_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_7.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_7.TabIndex = 26
        Me._lblDesc_7.Text = "Name"
        Me._lblDesc_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_6
        '
        Me._lblDesc_6.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_6, CType(6, Short))
        Me._lblDesc_6.Location = New System.Drawing.Point(6, 199)
        Me._lblDesc_6.Name = "_lblDesc_6"
        Me._lblDesc_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_6.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_6.TabIndex = 13
        Me._lblDesc_6.Text = "Polanyi Exponent"
        Me._lblDesc_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_5
        '
        Me._lblDesc_5.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_5, CType(5, Short))
        Me._lblDesc_5.Location = New System.Drawing.Point(6, 175)
        Me._lblDesc_5.Name = "_lblDesc_5"
        Me._lblDesc_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_5.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_5.TabIndex = 14
        Me._lblDesc_5.Text = "Polanyi B"
        Me._lblDesc_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_4
        '
        Me._lblDesc_4.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_4, CType(4, Short))
        Me._lblDesc_4.Location = New System.Drawing.Point(6, 151)
        Me._lblDesc_4.Name = "_lblDesc_4"
        Me._lblDesc_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_4.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_4.TabIndex = 15
        Me._lblDesc_4.Text = "Polanyi W0"
        Me._lblDesc_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_3
        '
        Me._lblDesc_3.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_3, CType(3, Short))
        Me._lblDesc_3.Location = New System.Drawing.Point(6, 127)
        Me._lblDesc_3.Name = "_lblDesc_3"
        Me._lblDesc_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_3.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_3.TabIndex = 16
        Me._lblDesc_3.Text = "Adsorbent Type"
        Me._lblDesc_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_2
        '
        Me._lblDesc_2.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_2, CType(2, Short))
        Me._lblDesc_2.Location = New System.Drawing.Point(6, 103)
        Me._lblDesc_2.Name = "_lblDesc_2"
        Me._lblDesc_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_2.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_2.TabIndex = 20
        Me._lblDesc_2.Text = "Particle Porosity"
        Me._lblDesc_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_1
        '
        Me._lblDesc_1.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_1, CType(1, Short))
        Me._lblDesc_1.Location = New System.Drawing.Point(6, 79)
        Me._lblDesc_1.Name = "_lblDesc_1"
        Me._lblDesc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_1.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_1.TabIndex = 21
        Me._lblDesc_1.Text = "Particle Radius"
        Me._lblDesc_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_0
        '
        Me._lblDesc_0.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_0, CType(0, Short))
        Me._lblDesc_0.Location = New System.Drawing.Point(6, 55)
        Me._lblDesc_0.Name = "_lblDesc_0"
        Me._lblDesc_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_0.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_0.TabIndex = 22
        Me._lblDesc_0.Text = "Apparent Density"
        Me._lblDesc_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblUnit_6
        '
        Me._lblUnit_6.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_6, CType(6, Short))
        Me._lblUnit_6.Location = New System.Drawing.Point(253, 199)
        Me._lblUnit_6.Name = "_lblUnit_6"
        Me._lblUnit_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_6.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_6.TabIndex = 10
        Me._lblUnit_6.Text = "-"
        Me._lblUnit_6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_5
        '
        Me._lblUnit_5.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_5, CType(5, Short))
        Me._lblUnit_5.Location = New System.Drawing.Point(253, 175)
        Me._lblUnit_5.Name = "_lblUnit_5"
        Me._lblUnit_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_5.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_5.TabIndex = 11
        Me._lblUnit_5.Text = "*"
        Me._lblUnit_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_4
        '
        Me._lblUnit_4.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_4, CType(4, Short))
        Me._lblUnit_4.Location = New System.Drawing.Point(253, 151)
        Me._lblUnit_4.Name = "_lblUnit_4"
        Me._lblUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_4.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_4.TabIndex = 12
        Me._lblUnit_4.Text = "cm3/g"
        Me._lblUnit_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_2
        '
        Me._lblUnit_2.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_2, CType(2, Short))
        Me._lblUnit_2.Location = New System.Drawing.Point(253, 103)
        Me._lblUnit_2.Name = "_lblUnit_2"
        Me._lblUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_2.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_2.TabIndex = 17
        Me._lblUnit_2.Text = "-"
        Me._lblUnit_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_1
        '
        Me._lblUnit_1.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_1, CType(1, Short))
        Me._lblUnit_1.Location = New System.Drawing.Point(253, 79)
        Me._lblUnit_1.Name = "_lblUnit_1"
        Me._lblUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_1.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_1.TabIndex = 18
        Me._lblUnit_1.Text = "cm"
        Me._lblUnit_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_0
        '
        Me._lblUnit_0.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_0, CType(0, Short))
        Me._lblUnit_0.Location = New System.Drawing.Point(253, 55)
        Me._lblUnit_0.Name = "_lblUnit_0"
        Me._lblUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_0.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_0.TabIndex = 19
        Me._lblUnit_0.Text = "g/cm3"
        Me._lblUnit_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtData
        '
        '
        '_txtData_7
        '
        Me._txtData_7.AcceptsReturn = True
        Me._txtData_7.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_7.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_7, CType(7, Short))
        Me._txtData_7.Location = New System.Drawing.Point(134, 26)
        Me._txtData_7.MaxLength = 0
        Me._txtData_7.Name = "_txtData_7"
        Me._txtData_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_7.Size = New System.Drawing.Size(176, 20)
        Me._txtData_7.TabIndex = 0
        Me._txtData_7.Text = "txtData(7)"
        Me._txtData_7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtData_6
        '
        Me._txtData_6.AcceptsReturn = True
        Me._txtData_6.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_6, CType(6, Short))
        Me._txtData_6.Location = New System.Drawing.Point(134, 197)
        Me._txtData_6.MaxLength = 0
        Me._txtData_6.Name = "_txtData_6"
        Me._txtData_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_6.Size = New System.Drawing.Size(113, 20)
        Me._txtData_6.TabIndex = 7
        Me._txtData_6.Text = "txtData(6)"
        Me._txtData_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtData_5
        '
        Me._txtData_5.AcceptsReturn = True
        Me._txtData_5.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_5, CType(5, Short))
        Me._txtData_5.Location = New System.Drawing.Point(134, 173)
        Me._txtData_5.MaxLength = 0
        Me._txtData_5.Name = "_txtData_5"
        Me._txtData_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_5.Size = New System.Drawing.Size(113, 20)
        Me._txtData_5.TabIndex = 6
        Me._txtData_5.Text = "txtData(5)"
        Me._txtData_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtData_4
        '
        Me._txtData_4.AcceptsReturn = True
        Me._txtData_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_4, CType(4, Short))
        Me._txtData_4.Location = New System.Drawing.Point(134, 149)
        Me._txtData_4.MaxLength = 0
        Me._txtData_4.Name = "_txtData_4"
        Me._txtData_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_4.Size = New System.Drawing.Size(113, 20)
        Me._txtData_4.TabIndex = 5
        Me._txtData_4.Text = "txtData(4)"
        Me._txtData_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtData_3
        '
        Me._txtData_3.AcceptsReturn = True
        Me._txtData_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_3, CType(3, Short))
        Me._txtData_3.Location = New System.Drawing.Point(134, 125)
        Me._txtData_3.MaxLength = 0
        Me._txtData_3.Name = "_txtData_3"
        Me._txtData_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_3.Size = New System.Drawing.Size(176, 20)
        Me._txtData_3.TabIndex = 4
        Me._txtData_3.Text = "txtData(3)"
        Me._txtData_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtData_2
        '
        Me._txtData_2.AcceptsReturn = True
        Me._txtData_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_2, CType(2, Short))
        Me._txtData_2.Location = New System.Drawing.Point(134, 101)
        Me._txtData_2.MaxLength = 0
        Me._txtData_2.Name = "_txtData_2"
        Me._txtData_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_2.Size = New System.Drawing.Size(113, 20)
        Me._txtData_2.TabIndex = 3
        Me._txtData_2.Text = "txtData(2)"
        Me._txtData_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtData_1
        '
        Me._txtData_1.AcceptsReturn = True
        Me._txtData_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_1, CType(1, Short))
        Me._txtData_1.Location = New System.Drawing.Point(134, 77)
        Me._txtData_1.MaxLength = 0
        Me._txtData_1.Name = "_txtData_1"
        Me._txtData_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_1.Size = New System.Drawing.Size(113, 20)
        Me._txtData_1.TabIndex = 2
        Me._txtData_1.Text = "txtData(1)"
        Me._txtData_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtData_0
        '
        Me._txtData_0.AcceptsReturn = True
        Me._txtData_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_0, CType(0, Short))
        Me._txtData_0.Location = New System.Drawing.Point(134, 53)
        Me._txtData_0.MaxLength = 0
        Me._txtData_0.Name = "_txtData_0"
        Me._txtData_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_0.Size = New System.Drawing.Size(113, 20)
        Me._txtData_0.TabIndex = 1
        Me._txtData_0.Text = "txtData(0)"
        Me._txtData_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_optPhase_1
        '
        Me._optPhase_1.AutoSize = True
        Me._optPhase_1.Checked = True
        Me._optPhase_1.Location = New System.Drawing.Point(46, 22)
        Me._optPhase_1.Name = "_optPhase_1"
        Me._optPhase_1.Size = New System.Drawing.Size(86, 18)
        Me._optPhase_1.TabIndex = 0
        Me._optPhase_1.TabStop = True
        Me._optPhase_1.Text = "&Liquid Phase"
        Me._optPhase_1.UseVisualStyleBackColor = True
        '
        '_optPhase_2
        '
        Me._optPhase_2.AutoSize = True
        Me._optPhase_2.Location = New System.Drawing.Point(182, 22)
        Me._optPhase_2.Name = "_optPhase_2"
        Me._optPhase_2.Size = New System.Drawing.Size(78, 18)
        Me._optPhase_2.TabIndex = 1
        Me._optPhase_2.Text = "&Gas Phase"
        Me._optPhase_2.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me._optPhase_2)
        Me.GroupBox2.Controls.Add(Me._optPhase_1)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(340, 54)
        Me.GroupBox2.TabIndex = 28
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Select Phase:"
        '
        'lblUnitB
        '
        Me.lblUnitB.BackColor = System.Drawing.Color.Transparent
        Me.lblUnitB.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUnitB.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUnitB.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnitB.Location = New System.Drawing.Point(43, 231)
        Me.lblUnitB.Name = "lblUnitB"
        Me.lblUnitB.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUnitB.Size = New System.Drawing.Size(267, 30)
        Me.lblUnitB.TabIndex = 9
        Me.lblUnitB.Text = "* in (mol/cal) ^(Polanyi Exponent)"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me._lblDesc_0)
        Me.GroupBox1.Controls.Add(Me._lblUnit_0)
        Me.GroupBox1.Controls.Add(Me._txtData_0)
        Me.GroupBox1.Controls.Add(Me._lblDesc_1)
        Me.GroupBox1.Controls.Add(Me._lblUnit_1)
        Me.GroupBox1.Controls.Add(Me._txtData_1)
        Me.GroupBox1.Controls.Add(Me._lblDesc_2)
        Me.GroupBox1.Controls.Add(Me._lblUnit_2)
        Me.GroupBox1.Controls.Add(Me._txtData_2)
        Me.GroupBox1.Controls.Add(Me._lblDesc_3)
        Me.GroupBox1.Controls.Add(Me._txtData_3)
        Me.GroupBox1.Controls.Add(Me._lblDesc_4)
        Me.GroupBox1.Controls.Add(Me._lblUnit_4)
        Me.GroupBox1.Controls.Add(Me._txtData_4)
        Me.GroupBox1.Controls.Add(Me._lblDesc_5)
        Me.GroupBox1.Controls.Add(Me._lblUnit_5)
        Me.GroupBox1.Controls.Add(Me._txtData_5)
        Me.GroupBox1.Controls.Add(Me._lblDesc_6)
        Me.GroupBox1.Controls.Add(Me._lblUnit_6)
        Me.GroupBox1.Controls.Add(Me._txtData_6)
        Me.GroupBox1.Controls.Add(Me._lblDesc_7)
        Me.GroupBox1.Controls.Add(Me._txtData_7)
        Me.GroupBox1.Controls.Add(Me.lblUnitB)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(340, 272)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Adsorbent Properties:"
        '
        '_cmdSaveCancel_0
        '
        Me._cmdSaveCancel_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdSaveCancel_0.Location = New System.Drawing.Point(5, 356)
        Me._cmdSaveCancel_0.Name = "_cmdSaveCancel_0"
        Me._cmdSaveCancel_0.Size = New System.Drawing.Size(100, 44)
        Me._cmdSaveCancel_0.TabIndex = 30
        Me._cmdSaveCancel_0.Text = "&Save"
        Me._cmdSaveCancel_0.UseVisualStyleBackColor = False
        '
        '_cmdSaveCancel_1
        '
        Me._cmdSaveCancel_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdSaveCancel_1.Location = New System.Drawing.Point(111, 356)
        Me._cmdSaveCancel_1.Name = "_cmdSaveCancel_1"
        Me._cmdSaveCancel_1.Size = New System.Drawing.Size(136, 44)
        Me._cmdSaveCancel_1.TabIndex = 31
        Me._cmdSaveCancel_1.Text = "Save &As New Record"
        Me._cmdSaveCancel_1.UseVisualStyleBackColor = False
        '
        '_cmdSaveCancel_2
        '
        Me._cmdSaveCancel_2.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdSaveCancel_2.Location = New System.Drawing.Point(252, 356)
        Me._cmdSaveCancel_2.Name = "_cmdSaveCancel_2"
        Me._cmdSaveCancel_2.Size = New System.Drawing.Size(100, 44)
        Me._cmdSaveCancel_2.TabIndex = 32
        Me._cmdSaveCancel_2.Text = "&Cancel"
        Me._cmdSaveCancel_2.UseVisualStyleBackColor = False
        '
        'frmEditCarbonData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(366, 410)
        Me.ControlBox = False
        Me.Controls.Add(Me._cmdSaveCancel_2)
        Me.Controls.Add(Me._cmdSaveCancel_1)
        Me.Controls.Add(Me._cmdSaveCancel_0)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(73, 152)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(382, 449)
        Me.Name = "frmEditCarbonData"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Editing an Adsorbent"
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents _optPhase_1 As RadioButton
    Friend WithEvents _optPhase_2 As RadioButton
    Friend WithEvents GroupBox2 As GroupBox
    Public WithEvents lblUnitB As Label
    Public WithEvents _txtData_7 As TextBox
    Public WithEvents _lblDesc_7 As Label
    Public WithEvents _txtData_6 As TextBox
    Public WithEvents _lblUnit_6 As Label
    Public WithEvents _lblDesc_6 As Label
    Public WithEvents _txtData_5 As TextBox
    Public WithEvents _lblUnit_5 As Label
    Public WithEvents _lblDesc_5 As Label
    Public WithEvents _txtData_4 As TextBox
    Public WithEvents _lblUnit_4 As Label
    Public WithEvents _lblDesc_4 As Label
    Public WithEvents _txtData_3 As TextBox
    Public WithEvents _lblDesc_3 As Label
    Public WithEvents _txtData_2 As TextBox
    Public WithEvents _lblUnit_2 As Label
    Public WithEvents _lblDesc_2 As Label
    Public WithEvents _txtData_1 As TextBox
    Public WithEvents _lblUnit_1 As Label
    Public WithEvents _lblDesc_1 As Label
    Public WithEvents _txtData_0 As TextBox
    Public WithEvents _lblUnit_0 As Label
    Public WithEvents _lblDesc_0 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents _cmdSaveCancel_0 As Button
    Friend WithEvents _cmdSaveCancel_1 As Button
    Friend WithEvents _cmdSaveCancel_2 As Button
#End Region
End Class