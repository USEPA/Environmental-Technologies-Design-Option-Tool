<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmKinetic
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
    Public WithEvents Picture1 As System.Windows.Forms.PictureBox
    Public WithEvents lblDP_OLD As System.Windows.Forms.Label
    Public WithEvents lblDS_OLD As System.Windows.Forms.Label
    Public WithEvents lblKF_OLD As System.Windows.Forms.Label
    Public WithEvents txtSPDFR As System.Windows.Forms.TextBox
    Public WithEvents txtTort As System.Windows.Forms.TextBox
    Public cmdCancelOK(2) As AxThreed.AxSSCommand
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents lblSPDFR As System.Windows.Forms.Label
    Public WithEvents lblTort As System.Windows.Forms.Label
    Public WithEvents lblTortCorrelation As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    '  Public WithEvents cmdCancelOK As SSCommandArray
    '  Public WithEvents fraKP As SSFrameArray
    Public WithEvents lblUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optDP As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    ' Public WithEvents optDP_old As SSOptionArray
    Public WithEvents optDS As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'Public WithEvents optDS_old As SSOptionArray
    Public WithEvents optKF As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'Public WithEvents optKF_old As SSOptionArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.txtSPDFR = New System.Windows.Forms.TextBox()
        Me.txtTort = New System.Windows.Forms.TextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.lblSPDFR = New System.Windows.Forms.Label()
        Me.lblTort = New System.Windows.Forms.Label()
        Me.lblTortCorrelation = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblUnit_0 = New System.Windows.Forms.Label()
        Me._lblUnit_1 = New System.Windows.Forms.Label()
        Me._lblUnit_2 = New System.Windows.Forms.Label()
        Me.optDP = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me._optDP_0 = New System.Windows.Forms.RadioButton()
        Me._optDP_1 = New System.Windows.Forms.RadioButton()
        Me.optDS = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me._optDS_0 = New System.Windows.Forms.RadioButton()
        Me._optDS_1 = New System.Windows.Forms.RadioButton()
        Me.optKF = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me._optKF_0 = New System.Windows.Forms.RadioButton()
        Me._optKF_1 = New System.Windows.Forms.RadioButton()
        Me.lblDP_OLD = New System.Windows.Forms.Label()
        Me.lblDS_OLD = New System.Windows.Forms.Label()
        Me.lblKF_OLD = New System.Windows.Forms.Label()
        Me._fraKP_0 = New System.Windows.Forms.GroupBox()
        Me.lblCorrelationKF = New System.Windows.Forms.Label()
        Me.txtKF = New System.Windows.Forms.TextBox()
        Me.lblKF = New System.Windows.Forms.TextBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me._fraKP_1 = New System.Windows.Forms.GroupBox()
        Me.lblDS = New System.Windows.Forms.TextBox()
        Me.txtDS = New System.Windows.Forms.TextBox()
        Me.lblCorrelationDS = New System.Windows.Forms.Label()
        Me.lblDP = New System.Windows.Forms.TextBox()
        Me.txtDP = New System.Windows.Forms.TextBox()
        Me.lblCorrelationDP = New System.Windows.Forms.Label()
        Me._fraKP_2 = New System.Windows.Forms.GroupBox()
        Me._cmdCancelOK_0 = New System.Windows.Forms.Button()
        Me._cmdCancelOK_1 = New System.Windows.Forms.Button()
        Me.chkTortuosity_Corr = New System.Windows.Forms.CheckBox()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabelDirty = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripLabelStatus = New System.Windows.Forms.ToolStripLabel()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optDP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optDS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optKF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._fraKP_0.SuspendLayout()
        Me._fraKP_1.SuspendLayout()
        Me._fraKP_2.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(568, 312)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 44
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'txtSPDFR
        '
        Me.txtSPDFR.AcceptsReturn = True
        Me.txtSPDFR.BackColor = System.Drawing.SystemColors.Window
        Me.txtSPDFR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSPDFR.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSPDFR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSPDFR.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSPDFR.Location = New System.Drawing.Point(104, 226)
        Me.txtSPDFR.MaxLength = 0
        Me.txtSPDFR.Name = "txtSPDFR"
        Me.txtSPDFR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSPDFR.Size = New System.Drawing.Size(73, 20)
        Me.txtSPDFR.TabIndex = 3
        Me.txtSPDFR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtTort
        '
        Me.txtTort.AcceptsReturn = True
        Me.txtTort.BackColor = System.Drawing.SystemColors.Window
        Me.txtTort.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTort.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTort.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTort.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTort.Location = New System.Drawing.Point(104, 254)
        Me.txtTort.MaxLength = 0
        Me.txtTort.Name = "txtTort"
        Me.txtTort.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTort.Size = New System.Drawing.Size(73, 20)
        Me.txtTort.TabIndex = 4
        Me.txtTort.Text = "txtTort"
        Me.txtTort.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.Color.Transparent
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(-3, 103)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(97, 24)
        Me._Label1_1.TabIndex = 9
        Me._Label1_1.Text = "Correlation"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.Color.Transparent
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(7, 75)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(85, 19)
        Me._Label1_0.TabIndex = 8
        Me._Label1_0.Text = "User Input"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSPDFR
        '
        Me.lblSPDFR.BackColor = System.Drawing.Color.Transparent
        Me.lblSPDFR.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSPDFR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSPDFR.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSPDFR.Location = New System.Drawing.Point(180, 232)
        Me.lblSPDFR.Name = "lblSPDFR"
        Me.lblSPDFR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSPDFR.Size = New System.Drawing.Size(333, 21)
        Me.lblSPDFR.TabIndex = 7
        Me.lblSPDFR.Text = "Surface To Pore Diffusion Flux Ratio (SPDFR)"
        '
        'lblTort
        '
        Me.lblTort.BackColor = System.Drawing.Color.Transparent
        Me.lblTort.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTort.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTort.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblTort.Location = New System.Drawing.Point(180, 260)
        Me.lblTort.Name = "lblTort"
        Me.lblTort.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTort.Size = New System.Drawing.Size(73, 17)
        Me.lblTort.TabIndex = 6
        Me.lblTort.Text = "Tortuosity"
        '
        'lblTortCorrelation
        '
        Me.lblTortCorrelation.BackColor = System.Drawing.Color.Transparent
        Me.lblTortCorrelation.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTortCorrelation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTortCorrelation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblTortCorrelation.Location = New System.Drawing.Point(104, 288)
        Me.lblTortCorrelation.Name = "lblTortCorrelation"
        Me.lblTortCorrelation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTortCorrelation.Size = New System.Drawing.Size(385, 31)
        Me.lblTortCorrelation.TabIndex = 5
        Me.lblTortCorrelation.Text = "Leave this label alone!"
        '
        '_lblUnit_0
        '
        Me._lblUnit_0.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_0, CType(0, Short))
        Me._lblUnit_0.Location = New System.Drawing.Point(30, 19)
        Me._lblUnit_0.Name = "_lblUnit_0"
        Me._lblUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_0.Size = New System.Drawing.Size(73, 17)
        Me._lblUnit_0.TabIndex = 41
        Me._lblUnit_0.Text = "cm/s"
        Me._lblUnit_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_1
        '
        Me._lblUnit_1.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_1, CType(1, Short))
        Me._lblUnit_1.Location = New System.Drawing.Point(30, 19)
        Me._lblUnit_1.Name = "_lblUnit_1"
        Me._lblUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_1.Size = New System.Drawing.Size(73, 17)
        Me._lblUnit_1.TabIndex = 19
        Me._lblUnit_1.Text = "cm2/s"
        Me._lblUnit_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_2
        '
        Me._lblUnit_2.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_2, CType(2, Short))
        Me._lblUnit_2.Location = New System.Drawing.Point(30, 19)
        Me._lblUnit_2.Name = "_lblUnit_2"
        Me._lblUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_2.Size = New System.Drawing.Size(73, 17)
        Me._lblUnit_2.TabIndex = 20
        Me._lblUnit_2.Text = "cm2/s"
        Me._lblUnit_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'optDP
        '
        '
        '_optDP_0
        '
        Me._optDP_0.AutoSize = True
        Me._optDP_0.Checked = True
        Me.optDP.SetIndex(Me._optDP_0, CType(0, Short))
        Me._optDP_0.Location = New System.Drawing.Point(10, 53)
        Me._optDP_0.Name = "_optDP_0"
        Me._optDP_0.Size = New System.Drawing.Size(14, 13)
        Me._optDP_0.TabIndex = 54
        Me._optDP_0.TabStop = True
        Me._optDP_0.UseVisualStyleBackColor = True
        '
        '_optDP_1
        '
        Me._optDP_1.AutoSize = True
        Me.optDP.SetIndex(Me._optDP_1, CType(1, Short))
        Me._optDP_1.Location = New System.Drawing.Point(10, 81)
        Me._optDP_1.Name = "_optDP_1"
        Me._optDP_1.Size = New System.Drawing.Size(14, 13)
        Me._optDP_1.TabIndex = 55
        Me._optDP_1.UseVisualStyleBackColor = True
        '
        'optDS
        '
        '
        '_optDS_0
        '
        Me._optDS_0.AutoSize = True
        Me._optDS_0.Checked = True
        Me.optDS.SetIndex(Me._optDS_0, CType(0, Short))
        Me._optDS_0.Location = New System.Drawing.Point(16, 53)
        Me._optDS_0.Name = "_optDS_0"
        Me._optDS_0.Size = New System.Drawing.Size(14, 13)
        Me._optDS_0.TabIndex = 53
        Me._optDS_0.TabStop = True
        Me._optDS_0.UseVisualStyleBackColor = True
        '
        '_optDS_1
        '
        Me._optDS_1.AutoSize = True
        Me.optDS.SetIndex(Me._optDS_1, CType(1, Short))
        Me._optDS_1.Location = New System.Drawing.Point(16, 81)
        Me._optDS_1.Name = "_optDS_1"
        Me._optDS_1.Size = New System.Drawing.Size(14, 13)
        Me._optDS_1.TabIndex = 54
        Me._optDS_1.UseVisualStyleBackColor = True
        '
        'optKF
        '
        '
        '_optKF_0
        '
        Me._optKF_0.BackColor = System.Drawing.SystemColors.Control
        Me._optKF_0.Checked = True
        Me._optKF_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optKF_0.Enabled = False
        Me._optKF_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._optKF_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optKF.SetIndex(Me._optKF_0, CType(0, Short))
        Me._optKF_0.Location = New System.Drawing.Point(7, 53)
        Me._optKF_0.Name = "_optKF_0"
        Me._optKF_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optKF_0.Size = New System.Drawing.Size(17, 17)
        Me._optKF_0.TabIndex = 43
        Me._optKF_0.TabStop = True
        Me._optKF_0.UseVisualStyleBackColor = False
        '
        '_optKF_1
        '
        Me._optKF_1.BackColor = System.Drawing.SystemColors.Control
        Me._optKF_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optKF_1.Enabled = False
        Me._optKF_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._optKF_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optKF.SetIndex(Me._optKF_1, CType(1, Short))
        Me._optKF_1.Location = New System.Drawing.Point(7, 81)
        Me._optKF_1.Name = "_optKF_1"
        Me._optKF_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optKF_1.Size = New System.Drawing.Size(17, 17)
        Me._optKF_1.TabIndex = 44
        Me._optKF_1.UseVisualStyleBackColor = False
        '
        'lblDP_OLD
        '
        Me.lblDP_OLD.BackColor = System.Drawing.SystemColors.Window
        Me.lblDP_OLD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDP_OLD.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDP_OLD.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDP_OLD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDP_OLD.Location = New System.Drawing.Point(36, 54)
        Me.lblDP_OLD.Name = "lblDP_OLD"
        Me.lblDP_OLD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDP_OLD.Size = New System.Drawing.Size(73, 19)
        Me.lblDP_OLD.TabIndex = 27
        Me.lblDP_OLD.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblDS_OLD
        '
        Me.lblDS_OLD.BackColor = System.Drawing.SystemColors.Window
        Me.lblDS_OLD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDS_OLD.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDS_OLD.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDS_OLD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDS_OLD.Location = New System.Drawing.Point(40, 46)
        Me.lblDS_OLD.Name = "lblDS_OLD"
        Me.lblDS_OLD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDS_OLD.Size = New System.Drawing.Size(73, 19)
        Me.lblDS_OLD.TabIndex = 26
        Me.lblDS_OLD.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblKF_OLD
        '
        Me.lblKF_OLD.BackColor = System.Drawing.SystemColors.Window
        Me.lblKF_OLD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblKF_OLD.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblKF_OLD.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKF_OLD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblKF_OLD.Location = New System.Drawing.Point(24, 28)
        Me.lblKF_OLD.Name = "lblKF_OLD"
        Me.lblKF_OLD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblKF_OLD.Size = New System.Drawing.Size(81, 19)
        Me.lblKF_OLD.TabIndex = 25
        Me.lblKF_OLD.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_fraKP_0
        '
        Me._fraKP_0.Controls.Add(Me.lblCorrelationKF)
        Me._fraKP_0.Controls.Add(Me.txtKF)
        Me._fraKP_0.Controls.Add(Me.lblKF)
        Me._fraKP_0.Controls.Add(Me._optKF_0)
        Me._fraKP_0.Controls.Add(Me._optKF_1)
        Me._fraKP_0.Controls.Add(Me._lblUnit_0)
        Me._fraKP_0.Location = New System.Drawing.Point(103, 22)
        Me._fraKP_0.Name = "_fraKP_0"
        Me._fraKP_0.Size = New System.Drawing.Size(140, 195)
        Me._fraKP_0.TabIndex = 48
        Me._fraKP_0.TabStop = False
        Me._fraKP_0.Text = "Film Diffusion"
        '
        'lblCorrelationKF
        '
        Me.lblCorrelationKF.BackColor = System.Drawing.Color.Transparent
        Me.lblCorrelationKF.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCorrelationKF.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCorrelationKF.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCorrelationKF.Location = New System.Drawing.Point(1, 116)
        Me.lblCorrelationKF.Name = "lblCorrelationKF"
        Me.lblCorrelationKF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCorrelationKF.Size = New System.Drawing.Size(131, 69)
        Me.lblCorrelationKF.TabIndex = 40
        Me.lblCorrelationKF.Text = "Wilke-Lee Modification of the Hirschfelder - Bird - Spotz method"
        '
        'txtKF
        '
        Me.txtKF.AcceptsReturn = True
        Me.txtKF.BackColor = System.Drawing.SystemColors.Window
        Me.txtKF.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtKF.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtKF.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtKF.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtKF.Location = New System.Drawing.Point(38, 46)
        Me.txtKF.MaxLength = 0
        Me.txtKF.Name = "txtKF"
        Me.txtKF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtKF.Size = New System.Drawing.Size(81, 20)
        Me.txtKF.TabIndex = 39
        Me.txtKF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblKF
        '
        Me.lblKF.AcceptsReturn = True
        Me.lblKF.BackColor = System.Drawing.SystemColors.Window
        Me.lblKF.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblKF.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.lblKF.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKF.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblKF.Location = New System.Drawing.Point(38, 75)
        Me.lblKF.MaxLength = 0
        Me.lblKF.Name = "lblKF"
        Me.lblKF.ReadOnly = True
        Me.lblKF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblKF.Size = New System.Drawing.Size(81, 20)
        Me.lblKF.TabIndex = 42
        Me.lblKF.TabStop = False
        Me.lblKF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_fraKP_1
        '
        Me._fraKP_1.Controls.Add(Me._optDS_1)
        Me._fraKP_1.Controls.Add(Me._optDS_0)
        Me._fraKP_1.Controls.Add(Me.lblDS)
        Me._fraKP_1.Controls.Add(Me.txtDS)
        Me._fraKP_1.Controls.Add(Me.lblCorrelationDS)
        Me._fraKP_1.Controls.Add(Me._lblUnit_1)
        Me._fraKP_1.Location = New System.Drawing.Point(255, 22)
        Me._fraKP_1.Name = "_fraKP_1"
        Me._fraKP_1.Size = New System.Drawing.Size(140, 195)
        Me._fraKP_1.TabIndex = 50
        Me._fraKP_1.TabStop = False
        Me._fraKP_1.Text = "Surface Diffusion"
        '
        'lblDS
        '
        Me.lblDS.AcceptsReturn = True
        Me.lblDS.BackColor = System.Drawing.SystemColors.Window
        Me.lblDS.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDS.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.lblDS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDS.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDS.Location = New System.Drawing.Point(52, 75)
        Me.lblDS.MaxLength = 0
        Me.lblDS.Name = "lblDS"
        Me.lblDS.ReadOnly = True
        Me.lblDS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDS.Size = New System.Drawing.Size(81, 20)
        Me.lblDS.TabIndex = 25
        Me.lblDS.TabStop = False
        Me.lblDS.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDS
        '
        Me.txtDS.AcceptsReturn = True
        Me.txtDS.BackColor = System.Drawing.SystemColors.Window
        Me.txtDS.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDS.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDS.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDS.Location = New System.Drawing.Point(51, 53)
        Me.txtDS.MaxLength = 0
        Me.txtDS.Name = "txtDS"
        Me.txtDS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDS.Size = New System.Drawing.Size(81, 20)
        Me.txtDS.TabIndex = 23
        Me.txtDS.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblCorrelationDS
        '
        Me.lblCorrelationDS.BackColor = System.Drawing.Color.Transparent
        Me.lblCorrelationDS.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCorrelationDS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCorrelationDS.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCorrelationDS.Location = New System.Drawing.Point(1, 116)
        Me.lblCorrelationDS.Name = "lblCorrelationDS"
        Me.lblCorrelationDS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCorrelationDS.Size = New System.Drawing.Size(131, 69)
        Me.lblCorrelationDS.TabIndex = 24
        Me.lblCorrelationDS.Text = "Wilke-Lee Modification of the Hirschfelder - Bird - Spotz method"
        '
        'lblDP
        '
        Me.lblDP.Location = New System.Drawing.Point(41, 75)
        Me.lblDP.Name = "lblDP"
        Me.lblDP.Size = New System.Drawing.Size(81, 20)
        Me.lblDP.TabIndex = 27
        '
        'txtDP
        '
        Me.txtDP.Location = New System.Drawing.Point(41, 46)
        Me.txtDP.Name = "txtDP"
        Me.txtDP.Size = New System.Drawing.Size(81, 20)
        Me.txtDP.TabIndex = 26
        '
        'lblCorrelationDP
        '
        Me.lblCorrelationDP.BackColor = System.Drawing.Color.Transparent
        Me.lblCorrelationDP.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCorrelationDP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCorrelationDP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCorrelationDP.Location = New System.Drawing.Point(1, 116)
        Me.lblCorrelationDP.Name = "lblCorrelationDP"
        Me.lblCorrelationDP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCorrelationDP.Size = New System.Drawing.Size(131, 69)
        Me.lblCorrelationDP.TabIndex = 28
        Me.lblCorrelationDP.Text = "Wilke-Lee Modification of the Hirschfelder - Bird - Spotz method"
        '
        '_fraKP_2
        '
        Me._fraKP_2.Controls.Add(Me._optDP_0)
        Me._fraKP_2.Controls.Add(Me._optDP_1)
        Me._fraKP_2.Controls.Add(Me.lblCorrelationDP)
        Me._fraKP_2.Controls.Add(Me.lblDP)
        Me._fraKP_2.Controls.Add(Me.txtDP)
        Me._fraKP_2.Controls.Add(Me._lblUnit_2)
        Me._fraKP_2.Location = New System.Drawing.Point(409, 22)
        Me._fraKP_2.Name = "_fraKP_2"
        Me._fraKP_2.Size = New System.Drawing.Size(140, 195)
        Me._fraKP_2.TabIndex = 52
        Me._fraKP_2.TabStop = False
        Me._fraKP_2.Text = "Pore Diffusion"
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(272, 381)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(100, 27)
        Me._cmdCancelOK_0.TabIndex = 53
        Me._cmdCancelOK_0.Text = "&Cancel"
        Me._cmdCancelOK_0.UseVisualStyleBackColor = False
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(141, 381)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(102, 27)
        Me._cmdCancelOK_1.TabIndex = 54
        Me._cmdCancelOK_1.Text = "&OK"
        Me._cmdCancelOK_1.UseVisualStyleBackColor = False
        '
        'chkTortuosity_Corr
        '
        Me.chkTortuosity_Corr.AutoSize = True
        Me.chkTortuosity_Corr.Location = New System.Drawing.Point(110, 322)
        Me.chkTortuosity_Corr.Name = "chkTortuosity_Corr"
        Me.chkTortuosity_Corr.Size = New System.Drawing.Size(234, 18)
        Me.chkTortuosity_Corr.TabIndex = 55
        Me.chkTortuosity_Corr.Text = "Use Pore Diffusion Correlation for &Totuosity"
        Me.chkTortuosity_Corr.UseVisualStyleBackColor = True
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabelDirty, Me.ToolStripLabelStatus})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 463)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(567, 25)
        Me.ToolStrip1.TabIndex = 56
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripLabelDirty
        '
        Me.ToolStripLabelDirty.AutoSize = False
        Me.ToolStripLabelDirty.Name = "ToolStripLabelDirty"
        Me.ToolStripLabelDirty.Size = New System.Drawing.Size(150, 22)
        Me.ToolStripLabelDirty.Text = "ToolStripLabelDirty"
        '
        'ToolStripLabelStatus
        '
        Me.ToolStripLabelStatus.AutoSize = False
        Me.ToolStripLabelStatus.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripLabelStatus.Name = "ToolStripLabelStatus"
        Me.ToolStripLabelStatus.Size = New System.Drawing.Size(350, 22)
        Me.ToolStripLabelStatus.Text = "ToolStripLabelStatus"
        '
        'frmKinetic
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(567, 488)
        Me.ControlBox = False
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.chkTortuosity_Corr)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Controls.Add(Me._fraKP_2)
        Me.Controls.Add(Me._fraKP_1)
        Me.Controls.Add(Me._fraKP_0)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.txtSPDFR)
        Me.Controls.Add(Me.txtTort)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me.lblSPDFR)
        Me.Controls.Add(Me.lblTort)
        Me.Controls.Add(Me.lblTortCorrelation)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(30, 154)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(583, 527)
        Me.Name = "frmKinetic"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Kinetic Parameters"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optDP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optDS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optKF, System.ComponentModel.ISupportInitialize).EndInit()
        Me._fraKP_0.ResumeLayout(False)
        Me._fraKP_0.PerformLayout()
        Me._fraKP_1.ResumeLayout(False)
        Me._fraKP_1.PerformLayout()
        Me._fraKP_2.ResumeLayout(False)
        Me._fraKP_2.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents _fraKP_0 As GroupBox
    Public WithEvents lblCorrelationKF As Label
    Public WithEvents txtKF As TextBox
    Public WithEvents lblKF As TextBox
    Public WithEvents _optKF_0 As RadioButton
    Public WithEvents _optKF_1 As RadioButton
    Public WithEvents _lblUnit_0 As Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents _fraKP_1 As GroupBox
    Public WithEvents lblDS As TextBox
    Public WithEvents txtDS As TextBox
    Public WithEvents lblCorrelationDS As Label
    Public WithEvents _lblUnit_1 As Label
    Public WithEvents lblDP As TextBox
    Public WithEvents txtDP As TextBox
    Public WithEvents lblCorrelationDP As Label
    Friend WithEvents _fraKP_2 As GroupBox

    Public WithEvents _lblUnit_2 As Label
    Friend WithEvents _optDS_1 As RadioButton
    Friend WithEvents _optDS_0 As RadioButton
    Friend WithEvents _optDP_1 As RadioButton
    Friend WithEvents _optDP_0 As RadioButton
    Friend WithEvents _cmdCancelOK_0 As Button
    Friend WithEvents _cmdCancelOK_1 As Button
    Friend WithEvents chkTortuosity_Corr As CheckBox
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripLabelDirty As ToolStripLabel
    Friend WithEvents ToolStripLabelStatus As ToolStripLabel


#End Region
End Class