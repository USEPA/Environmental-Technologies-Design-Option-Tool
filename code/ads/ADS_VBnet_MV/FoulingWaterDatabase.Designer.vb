<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFoulingWaterDatabase
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
	Public WithEvents lstCorrelations As System.Windows.Forms.ListBox
    Public WithEvents SSFrame1 As AxThreed.AxSSFrame
    Public WithEvents _txtCoeff_4 As System.Windows.Forms.TextBox
    Public WithEvents _txtCoeff_3 As System.Windows.Forms.TextBox
    Public WithEvents _txtCoeff_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtCoeff_1 As System.Windows.Forms.TextBox
    Public WithEvents txtName As System.Windows.Forms.TextBox
    Public WithEvents _lblDesc_1 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_2 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_3 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_4 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_0 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_1 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_2 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_3 As System.Windows.Forms.Label
    Public WithEvents SSFrame2 As AxThreed.AxSSFrame
    Public WithEvents _cmdCancelOK_1 As AxThreed.AxSSCommand
    Public WithEvents _cmdCancelOK_0 As AxThreed.AxSSCommand
    '   Public WithEvents cmdCancelOK As SSCommandArray
    '   Public WithEvents cmdRecord As SSCommandArray
    Public WithEvents lblDesc As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtCoeff As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFoulingWaterDatabase))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblDesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtCoeff = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.lstCorrelations = New System.Windows.Forms.ListBox()
        Me._txtCoeff_4 = New System.Windows.Forms.TextBox()
        Me._txtCoeff_3 = New System.Windows.Forms.TextBox()
        Me._txtCoeff_2 = New System.Windows.Forms.TextBox()
        Me._txtCoeff_1 = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me._lblDesc_1 = New System.Windows.Forms.Label()
        Me._lblDesc_2 = New System.Windows.Forms.Label()
        Me._lblDesc_3 = New System.Windows.Forms.Label()
        Me._lblDesc_4 = New System.Windows.Forms.Label()
        Me._lblUnit_0 = New System.Windows.Forms.Label()
        Me._lblUnit_1 = New System.Windows.Forms.Label()
        Me._lblUnit_2 = New System.Windows.Forms.Label()
        Me._lblUnit_3 = New System.Windows.Forms.Label()
        Me.SSFrame1 = New AxThreed.AxSSFrame()
        Me.SSFrame2 = New AxThreed.AxSSFrame()
        Me._cmdCancelOK_1 = New AxThreed.AxSSCommand()
        Me._cmdCancelOK_0 = New AxThreed.AxSSCommand()
        Me._cmdRecord_0 = New AxThreed.AxSSCommand()
        Me._cmdRecord_1 = New AxThreed.AxSSCommand()
        Me._cmdRecord_2 = New AxThreed.AxSSCommand()
        Me._cmdRecord_3 = New AxThreed.AxSSCommand()
        Me._cmdRecord_4 = New AxThreed.AxSSCommand()
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCoeff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSFrame1.SuspendLayout()
        CType(Me.SSFrame2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSFrame2.SuspendLayout()
        CType(Me._cmdCancelOK_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdCancelOK_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdRecord_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdRecord_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdRecord_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdRecord_3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdRecord_4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCoeff
        '
        '
        'lstCorrelations
        '
        Me.lstCorrelations.BackColor = System.Drawing.SystemColors.Window
        Me.lstCorrelations.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstCorrelations.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCorrelations.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCorrelations.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCorrelations.ItemHeight = 16
        Me.lstCorrelations.Location = New System.Drawing.Point(8, 18)
        Me.lstCorrelations.Name = "lstCorrelations"
        Me.lstCorrelations.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCorrelations.Size = New System.Drawing.Size(386, 130)
        Me.lstCorrelations.TabIndex = 2
        Me.lstCorrelations.TabStop = False
        '
        '_txtCoeff_4
        '
        Me._txtCoeff_4.AcceptsReturn = True
        Me._txtCoeff_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtCoeff_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtCoeff_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCoeff_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtCoeff_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtCoeff_4.Location = New System.Drawing.Point(132, 130)
        Me._txtCoeff_4.MaxLength = 0
        Me._txtCoeff_4.Name = "_txtCoeff_4"
        Me._txtCoeff_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_4.Size = New System.Drawing.Size(106, 23)
        Me._txtCoeff_4.TabIndex = 10
        Me._txtCoeff_4.Text = "txtCoeff(4)"
        Me.txtCoeff.SetIndex(_txtCoeff_4, CType(4, Short))
        Me._txtCoeff_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtCoeff_3
        '
        Me._txtCoeff_3.AcceptsReturn = True
        Me._txtCoeff_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtCoeff_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtCoeff_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCoeff_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtCoeff_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtCoeff_3.Location = New System.Drawing.Point(132, 106)
        Me._txtCoeff_3.MaxLength = 0
        Me._txtCoeff_3.Name = "_txtCoeff_3"
        Me._txtCoeff_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_3.Size = New System.Drawing.Size(106, 23)
        Me._txtCoeff_3.TabIndex = 9
        Me._txtCoeff_3.Text = "txtCoeff(3)"
        Me.txtCoeff.SetIndex(_txtCoeff_3, CType(3, Short))

        Me._txtCoeff_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtCoeff_2
        '
        Me._txtCoeff_2.AcceptsReturn = True
        Me._txtCoeff_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtCoeff_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtCoeff_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCoeff_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtCoeff_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtCoeff_2.Location = New System.Drawing.Point(132, 82)
        Me._txtCoeff_2.MaxLength = 0
        Me._txtCoeff_2.Name = "_txtCoeff_2"
        Me._txtCoeff_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_2.Size = New System.Drawing.Size(106, 23)
        Me._txtCoeff_2.TabIndex = 8
        Me._txtCoeff_2.Text = "txtCoeff(2)"
        Me.txtCoeff.SetIndex(_txtCoeff_2, CType(2, Short))

        Me._txtCoeff_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtCoeff_1
        '
        Me._txtCoeff_1.AcceptsReturn = True
        Me._txtCoeff_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtCoeff_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtCoeff_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCoeff_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtCoeff_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtCoeff_1.Location = New System.Drawing.Point(132, 58)
        Me._txtCoeff_1.MaxLength = 0
        Me._txtCoeff_1.Name = "_txtCoeff_1"
        Me._txtCoeff_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_1.Size = New System.Drawing.Size(106, 23)
        Me._txtCoeff_1.TabIndex = 7
        Me._txtCoeff_1.Text = "txtCoeff(1)"
        Me.txtCoeff.SetIndex(_txtCoeff_1, CType(1, Short))

        Me._txtCoeff_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtName
        '
        Me.txtName.AcceptsReturn = True
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtName.Location = New System.Drawing.Point(8, 22)
        Me.txtName.MaxLength = 0
        Me.txtName.Name = "txtName"
        Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtName.Size = New System.Drawing.Size(378, 23)
        Me.txtName.TabIndex = 3
        Me.txtName.Text = "txtName"
        Me.txtName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_lblDesc_1
        '
        Me._lblDesc_1.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblDesc_1.Location = New System.Drawing.Point(48, 60)
        Me._lblDesc_1.Name = "_lblDesc_1"
        Me._lblDesc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_1.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_1.TabIndex = 18
        Me._lblDesc_1.Text = "K1"
        Me.lblDesc.SetIndex(_lblDesc_1, CType(1, Short))

        Me._lblDesc_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_2
        '
        Me._lblDesc_2.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblDesc_2.Location = New System.Drawing.Point(48, 84)
        Me._lblDesc_2.Name = "_lblDesc_2"
        Me._lblDesc_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_2.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_2.TabIndex = 17
        Me._lblDesc_2.Text = "K2"
        Me.lblDesc.SetIndex(_lblDesc_2, CType(2, Short))
        Me._lblDesc_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_3
        '
        Me._lblDesc_3.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblDesc_3.Location = New System.Drawing.Point(48, 108)
        Me._lblDesc_3.Name = "_lblDesc_3"
        Me._lblDesc_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_3.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_3.TabIndex = 16
        Me._lblDesc_3.Text = "K3"
        Me.lblDesc.SetIndex(_lblDesc_3, CType(3, Short))

        Me._lblDesc_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_4
        '
        Me._lblDesc_4.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblDesc_4.Location = New System.Drawing.Point(48, 132)
        Me._lblDesc_4.Name = "_lblDesc_4"
        Me._lblDesc_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_4.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_4.TabIndex = 15
        Me._lblDesc_4.Text = "K4"
        Me.lblDesc.SetIndex(_lblDesc_4, CType(4, Short))

        Me._lblDesc_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblUnit_0
        '
        Me._lblUnit_0.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblUnit_0.Location = New System.Drawing.Point(252, 60)
        Me._lblUnit_0.Name = "_lblUnit_0"
        Me._lblUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_0.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_0.TabIndex = 14
        Me._lblUnit_0.Text = "-"
        Me.lblUnit.SetIndex(_lblUnit_0, CType(0, Short))

        Me._lblUnit_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_1
        '
        Me._lblUnit_1.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblUnit_1.Location = New System.Drawing.Point(252, 84)
        Me._lblUnit_1.Name = "_lblUnit_1"
        Me._lblUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_1.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_1.TabIndex = 13
        Me._lblUnit_1.Text = "1/min"
        Me.lblUnit.SetIndex(_lblUnit_1, CType(1, Short))

        Me._lblUnit_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_2
        '
        Me._lblUnit_2.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblUnit_2.Location = New System.Drawing.Point(252, 108)
        Me._lblUnit_2.Name = "_lblUnit_2"
        Me._lblUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_2.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_2.TabIndex = 12
        Me._lblUnit_2.Text = "-"
        Me.lblUnit.SetIndex(_lblUnit_2, CType(2, Short))

        Me._lblUnit_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_3
        '
        Me._lblUnit_3.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lblUnit_3.Location = New System.Drawing.Point(252, 132)
        Me._lblUnit_3.Name = "_lblUnit_3"
        Me._lblUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_3.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_3.TabIndex = 11
        Me._lblUnit_3.Text = "1/min"
        Me.lblUnit.SetIndex(_lblUnit_3, CType(3, Short))

        Me._lblUnit_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'SSFrame1
        '
        Me.SSFrame1.Controls.Add(Me.lstCorrelations)
        Me.SSFrame1.Location = New System.Drawing.Point(8, 8)
        Me.SSFrame1.Name = "SSFrame1"
        Me.SSFrame1.OcxState = CType(resources.GetObject("SSFrame1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame1.Size = New System.Drawing.Size(408, 167)
        Me.SSFrame1.TabIndex = 0
        '
        'SSFrame2
        '
        Me.SSFrame2.Controls.Add(Me.txtName)
        Me.SSFrame2.Controls.Add(Me._txtCoeff_1)
        Me.SSFrame2.Controls.Add(Me._txtCoeff_2)
        Me.SSFrame2.Controls.Add(Me._txtCoeff_3)
        Me.SSFrame2.Controls.Add(Me._txtCoeff_4)
        Me.SSFrame2.Controls.Add(Me._lblDesc_1)
        Me.SSFrame2.Controls.Add(Me._lblDesc_2)
        Me.SSFrame2.Controls.Add(Me._lblDesc_3)
        Me.SSFrame2.Controls.Add(Me._lblDesc_4)
        Me.SSFrame2.Controls.Add(Me._lblUnit_0)
        Me.SSFrame2.Controls.Add(Me._lblUnit_1)
        Me.SSFrame2.Controls.Add(Me._lblUnit_2)
        Me.SSFrame2.Controls.Add(Me._lblUnit_3)
        Me.SSFrame2.Location = New System.Drawing.Point(8, 181)
        Me.SSFrame2.Name = "SSFrame2"
        Me.SSFrame2.OcxState = CType(resources.GetObject("SSFrame2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame2.Size = New System.Drawing.Size(408, 245)
        Me.SSFrame2.TabIndex = 1
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(72, 432)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.OcxState = CType(resources.GetObject("_cmdCancelOK_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(100, 50)
        Me._cmdCancelOK_1.TabIndex = 19
        Me._cmdCancelOK_1.TabStop = False
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(245, 432)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.OcxState = CType(resources.GetObject("_cmdCancelOK_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(100, 50)
        Me._cmdCancelOK_0.TabIndex = 20
        Me._cmdCancelOK_0.TabStop = False
        '
        '_cmdRecord_0
        '
        Me._cmdRecord_0.Location = New System.Drawing.Point(22, 355)
        Me._cmdRecord_0.Name = "_cmdRecord_0"
        Me._cmdRecord_0.OcxState = CType(resources.GetObject("_cmdRecord_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdRecord_0.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_0.TabIndex = 23
        Me._cmdRecord_0.TabStop = False
        '
        '_cmdRecord_1
        '
        Me._cmdRecord_1.Location = New System.Drawing.Point(97, 355)
        Me._cmdRecord_1.Name = "_cmdRecord_1"
        Me._cmdRecord_1.OcxState = CType(resources.GetObject("_cmdRecord_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdRecord_1.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_1.TabIndex = 24
        Me._cmdRecord_1.TabStop = False
        '
        '_cmdRecord_2
        '
        Me._cmdRecord_2.Location = New System.Drawing.Point(171, 355)
        Me._cmdRecord_2.Name = "_cmdRecord_2"
        Me._cmdRecord_2.OcxState = CType(resources.GetObject("_cmdRecord_2.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdRecord_2.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_2.TabIndex = 25
        Me._cmdRecord_2.TabStop = False
        '
        '_cmdRecord_3
        '
        Me._cmdRecord_3.Location = New System.Drawing.Point(245, 355)
        Me._cmdRecord_3.Name = "_cmdRecord_3"
        Me._cmdRecord_3.OcxState = CType(resources.GetObject("_cmdRecord_3.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdRecord_3.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_3.TabIndex = 26
        Me._cmdRecord_3.TabStop = False
        '
        '_cmdRecord_4
        '
        Me._cmdRecord_4.Location = New System.Drawing.Point(319, 355)
        Me._cmdRecord_4.Name = "_cmdRecord_4"
        Me._cmdRecord_4.OcxState = CType(resources.GetObject("_cmdRecord_4.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdRecord_4.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_4.TabIndex = 27
        Me._cmdRecord_4.TabStop = False
        '
        'frmFoulingWaterDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(450, 504)
        Me.ControlBox = False
        Me.Controls.Add(Me._cmdRecord_0)
        Me.Controls.Add(Me._cmdRecord_1)
        Me.Controls.Add(Me._cmdRecord_2)
        Me.Controls.Add(Me._cmdRecord_3)
        Me.Controls.Add(Me._cmdRecord_4)
        Me.Controls.Add(Me.SSFrame1)
        Me.Controls.Add(Me.SSFrame2)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(114, 177)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFoulingWaterDatabase"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Water Fouling Correlation Database"
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCoeff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSFrame1.ResumeLayout(False)
        CType(Me.SSFrame2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSFrame2.ResumeLayout(False)
        Me.SSFrame2.PerformLayout()
        CType(Me._cmdCancelOK_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdCancelOK_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdRecord_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdRecord_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdRecord_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdRecord_3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdRecord_4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Public WithEvents _cmdRecord_0 As AxThreed.AxSSCommand
    Public WithEvents _cmdRecord_1 As AxThreed.AxSSCommand
    Public WithEvents _cmdRecord_2 As AxThreed.AxSSCommand
    Public WithEvents _cmdRecord_3 As AxThreed.AxSSCommand
    Public WithEvents _cmdRecord_4 As AxThreed.AxSSCommand
#End Region
End Class