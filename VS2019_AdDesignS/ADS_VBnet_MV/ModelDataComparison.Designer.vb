<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmModelDataComparison
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
    Public WithEvents cboCUnits As System.Windows.Forms.ComboBox
    Public WithEvents cboGraphType As System.Windows.Forms.ComboBox
    Public WithEvents cboCompo As System.Windows.Forms.ComboBox
    Public WithEvents cboGrid As System.Windows.Forms.ComboBox
    Public WithEvents cboTUnits As System.Windows.Forms.ComboBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.cboCUnits = New System.Windows.Forms.ComboBox()
        Me.cboGraphType = New System.Windows.Forms.ComboBox()
        Me.cboCompo = New System.Windows.Forms.ComboBox()
        Me.cboGrid = New System.Windows.Forms.ComboBox()
        Me.cboTUnits = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(735, 27)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 13
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'cboCUnits
        '
        Me.cboCUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboCUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCUnits.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCUnits.Location = New System.Drawing.Point(6, 19)
        Me.cboCUnits.Name = "cboCUnits"
        Me.cboCUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCUnits.Size = New System.Drawing.Size(76, 22)
        Me.cboCUnits.TabIndex = 8
        '
        'cboGraphType
        '
        Me.cboGraphType.BackColor = System.Drawing.SystemColors.Window
        Me.cboGraphType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboGraphType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGraphType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboGraphType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboGraphType.Location = New System.Drawing.Point(490, 4)
        Me.cboGraphType.Name = "cboGraphType"
        Me.cboGraphType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboGraphType.Size = New System.Drawing.Size(89, 22)
        Me.cboGraphType.TabIndex = 2
        '
        'cboCompo
        '
        Me.cboCompo.BackColor = System.Drawing.SystemColors.Window
        Me.cboCompo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCompo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCompo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCompo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCompo.Location = New System.Drawing.Point(157, 4)
        Me.cboCompo.Name = "cboCompo"
        Me.cboCompo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCompo.Size = New System.Drawing.Size(213, 22)
        Me.cboCompo.TabIndex = 1
        '
        'cboGrid
        '
        Me.cboGrid.BackColor = System.Drawing.SystemColors.Window
        Me.cboGrid.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboGrid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGrid.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboGrid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboGrid.Location = New System.Drawing.Point(490, 28)
        Me.cboGrid.Name = "cboGrid"
        Me.cboGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboGrid.Size = New System.Drawing.Size(89, 22)
        Me.cboGrid.TabIndex = 0
        '
        'cboTUnits
        '
        Me.cboTUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboTUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTUnits.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTUnits.Location = New System.Drawing.Point(6, 19)
        Me.cboTUnits.Name = "cboTUnits"
        Me.cboTUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTUnits.Size = New System.Drawing.Size(78, 22)
        Me.cboTUnits.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label3.Location = New System.Drawing.Point(-2, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(157, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Select a component:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(378, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(106, 24)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Plot Patterns:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label5.Location = New System.Drawing.Point(395, 32)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(89, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Grid Style:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboCUnits)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 43)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(95, 51)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "C Units"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboTUnits)
        Me.GroupBox2.Location = New System.Drawing.Point(125, 43)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(95, 51)
        Me.GroupBox2.TabIndex = 15
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "T Units"
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdClose.Location = New System.Drawing.Point(586, 3)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(100, 22)
        Me.cmdClose.TabIndex = 16
        Me.cmdClose.Text = "E&xit"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'Chart1
        '
        ChartArea1.AxisX.TitleFont = New System.Drawing.Font("Garamond", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.AxisY.TitleFont = New System.Drawing.Font("Garamond", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Font = New System.Drawing.Font("Agency FB", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Legend1.IsTextAutoFit = False
        Legend1.Name = "Legend1"
        Legend1.TextWrapThreshold = 13
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(12, 100)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(714, 473)
        Me.Chart1.TabIndex = 18
        Me.Chart1.Text = "Chart1"
        '
        'frmModelDataComparison
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(740, 585)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.cboGraphType)
        Me.Controls.Add(Me.cboCompo)
        Me.Controls.Add(Me.cboGrid)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(23, 123)
        Me.MinimumSize = New System.Drawing.Size(756, 624)
        Me.Name = "frmModelDataComparison"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Data Comparison"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents cmdClose As Button
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
#End Region
End Class