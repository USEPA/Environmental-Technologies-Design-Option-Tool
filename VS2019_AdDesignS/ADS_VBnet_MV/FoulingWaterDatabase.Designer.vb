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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblDesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblDesc_1 = New System.Windows.Forms.Label()
        Me._lblDesc_2 = New System.Windows.Forms.Label()
        Me._lblDesc_3 = New System.Windows.Forms.Label()
        Me._lblDesc_4 = New System.Windows.Forms.Label()
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblUnit_0 = New System.Windows.Forms.Label()
        Me._lblUnit_1 = New System.Windows.Forms.Label()
        Me._lblUnit_2 = New System.Windows.Forms.Label()
        Me._lblUnit_3 = New System.Windows.Forms.Label()
        Me.txtCoeff = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me._txtCoeff_4 = New System.Windows.Forms.TextBox()
        Me._txtCoeff_3 = New System.Windows.Forms.TextBox()
        Me._txtCoeff_2 = New System.Windows.Forms.TextBox()
        Me._txtCoeff_1 = New System.Windows.Forms.TextBox()
        Me.lstCorrelations = New System.Windows.Forms.ListBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me._cmdRecord_0 = New System.Windows.Forms.Button()
        Me._cmdRecord_1 = New System.Windows.Forms.Button()
        Me._cmdRecord_2 = New System.Windows.Forms.Button()
        Me._cmdRecord_3 = New System.Windows.Forms.Button()
        Me._cmdRecord_4 = New System.Windows.Forms.Button()
        Me._cmdCancelOK_1 = New System.Windows.Forms.Button()
        Me._cmdCancelOK_0 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCoeff, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        '_lblDesc_1
        '
        Me._lblDesc_1.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_1, CType(1, Short))
        Me._lblDesc_1.Location = New System.Drawing.Point(54, 59)
        Me._lblDesc_1.Name = "_lblDesc_1"
        Me._lblDesc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_1.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_1.TabIndex = 18
        Me._lblDesc_1.Text = "K1"
        Me._lblDesc_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_2
        '
        Me._lblDesc_2.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_2, CType(2, Short))
        Me._lblDesc_2.Location = New System.Drawing.Point(54, 83)
        Me._lblDesc_2.Name = "_lblDesc_2"
        Me._lblDesc_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_2.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_2.TabIndex = 17
        Me._lblDesc_2.Text = "K2"
        Me._lblDesc_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_3
        '
        Me._lblDesc_3.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_3, CType(3, Short))
        Me._lblDesc_3.Location = New System.Drawing.Point(54, 107)
        Me._lblDesc_3.Name = "_lblDesc_3"
        Me._lblDesc_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_3.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_3.TabIndex = 16
        Me._lblDesc_3.Text = "K3"
        Me._lblDesc_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_4
        '
        Me._lblDesc_4.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_4, CType(4, Short))
        Me._lblDesc_4.Location = New System.Drawing.Point(54, 131)
        Me._lblDesc_4.Name = "_lblDesc_4"
        Me._lblDesc_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_4.Size = New System.Drawing.Size(73, 17)
        Me._lblDesc_4.TabIndex = 15
        Me._lblDesc_4.Text = "K4"
        Me._lblDesc_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblUnit_0
        '
        Me._lblUnit_0.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_0, CType(0, Short))
        Me._lblUnit_0.Location = New System.Drawing.Point(258, 59)
        Me._lblUnit_0.Name = "_lblUnit_0"
        Me._lblUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_0.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_0.TabIndex = 14
        Me._lblUnit_0.Text = "-"
        Me._lblUnit_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_1
        '
        Me._lblUnit_1.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_1, CType(1, Short))
        Me._lblUnit_1.Location = New System.Drawing.Point(258, 83)
        Me._lblUnit_1.Name = "_lblUnit_1"
        Me._lblUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_1.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_1.TabIndex = 13
        Me._lblUnit_1.Text = "1/min"
        Me._lblUnit_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_2
        '
        Me._lblUnit_2.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_2, CType(2, Short))
        Me._lblUnit_2.Location = New System.Drawing.Point(258, 107)
        Me._lblUnit_2.Name = "_lblUnit_2"
        Me._lblUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_2.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_2.TabIndex = 12
        Me._lblUnit_2.Text = "-"
        Me._lblUnit_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_3
        '
        Me._lblUnit_3.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_3, CType(3, Short))
        Me._lblUnit_3.Location = New System.Drawing.Point(258, 131)
        Me._lblUnit_3.Name = "_lblUnit_3"
        Me._lblUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_3.Size = New System.Drawing.Size(36, 17)
        Me._lblUnit_3.TabIndex = 11
        Me._lblUnit_3.Text = "1/min"
        Me._lblUnit_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCoeff
        '
        '
        '_txtCoeff_4
        '
        Me._txtCoeff_4.AcceptsReturn = True
        Me._txtCoeff_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtCoeff_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtCoeff_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCoeff_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtCoeff_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoeff.SetIndex(Me._txtCoeff_4, CType(4, Short))
        Me._txtCoeff_4.Location = New System.Drawing.Point(138, 129)
        Me._txtCoeff_4.MaxLength = 0
        Me._txtCoeff_4.Name = "_txtCoeff_4"
        Me._txtCoeff_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_4.Size = New System.Drawing.Size(106, 20)
        Me._txtCoeff_4.TabIndex = 10
        Me._txtCoeff_4.Text = "txtCoeff(4)"
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
        Me.txtCoeff.SetIndex(Me._txtCoeff_3, CType(3, Short))
        Me._txtCoeff_3.Location = New System.Drawing.Point(138, 105)
        Me._txtCoeff_3.MaxLength = 0
        Me._txtCoeff_3.Name = "_txtCoeff_3"
        Me._txtCoeff_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_3.Size = New System.Drawing.Size(106, 20)
        Me._txtCoeff_3.TabIndex = 9
        Me._txtCoeff_3.Text = "txtCoeff(3)"
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
        Me.txtCoeff.SetIndex(Me._txtCoeff_2, CType(2, Short))
        Me._txtCoeff_2.Location = New System.Drawing.Point(138, 81)
        Me._txtCoeff_2.MaxLength = 0
        Me._txtCoeff_2.Name = "_txtCoeff_2"
        Me._txtCoeff_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_2.Size = New System.Drawing.Size(106, 20)
        Me._txtCoeff_2.TabIndex = 8
        Me._txtCoeff_2.Text = "txtCoeff(2)"
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
        Me.txtCoeff.SetIndex(Me._txtCoeff_1, CType(1, Short))
        Me._txtCoeff_1.Location = New System.Drawing.Point(138, 57)
        Me._txtCoeff_1.MaxLength = 0
        Me._txtCoeff_1.Name = "_txtCoeff_1"
        Me._txtCoeff_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_1.Size = New System.Drawing.Size(106, 20)
        Me._txtCoeff_1.TabIndex = 7
        Me._txtCoeff_1.Text = "txtCoeff(1)"
        Me._txtCoeff_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lstCorrelations
        '
        Me.lstCorrelations.BackColor = System.Drawing.SystemColors.Window
        Me.lstCorrelations.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstCorrelations.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCorrelations.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCorrelations.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCorrelations.ItemHeight = 14
        Me.lstCorrelations.Location = New System.Drawing.Point(6, 19)
        Me.lstCorrelations.Name = "lstCorrelations"
        Me.lstCorrelations.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCorrelations.Size = New System.Drawing.Size(386, 128)
        Me.lstCorrelations.TabIndex = 2
        Me.lstCorrelations.TabStop = False
        '
        'txtName
        '
        Me.txtName.AcceptsReturn = True
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtName.Location = New System.Drawing.Point(14, 21)
        Me.txtName.MaxLength = 0
        Me.txtName.Name = "txtName"
        Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtName.Size = New System.Drawing.Size(378, 20)
        Me.txtName.TabIndex = 3
        Me.txtName.Text = "txtName"
        Me.txtName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_cmdRecord_0
        '
        Me._cmdRecord_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdRecord_0.Location = New System.Drawing.Point(14, 175)
        Me._cmdRecord_0.Name = "_cmdRecord_0"
        Me._cmdRecord_0.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_0.TabIndex = 28
        Me._cmdRecord_0.Text = "&New"
        Me._cmdRecord_0.UseVisualStyleBackColor = False
        '
        '_cmdRecord_1
        '
        Me._cmdRecord_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdRecord_1.Location = New System.Drawing.Point(89, 175)
        Me._cmdRecord_1.Name = "_cmdRecord_1"
        Me._cmdRecord_1.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_1.TabIndex = 29
        Me._cmdRecord_1.Text = "&Edit"
        Me._cmdRecord_1.UseVisualStyleBackColor = False
        '
        '_cmdRecord_2
        '
        Me._cmdRecord_2.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdRecord_2.Location = New System.Drawing.Point(163, 175)
        Me._cmdRecord_2.Name = "_cmdRecord_2"
        Me._cmdRecord_2.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_2.TabIndex = 30
        Me._cmdRecord_2.Text = "&Delete"
        Me._cmdRecord_2.UseVisualStyleBackColor = False
        '
        '_cmdRecord_3
        '
        Me._cmdRecord_3.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdRecord_3.Location = New System.Drawing.Point(237, 175)
        Me._cmdRecord_3.Name = "_cmdRecord_3"
        Me._cmdRecord_3.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_3.TabIndex = 31
        Me._cmdRecord_3.Text = "&Save"
        Me._cmdRecord_3.UseVisualStyleBackColor = False
        '
        '_cmdRecord_4
        '
        Me._cmdRecord_4.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdRecord_4.Location = New System.Drawing.Point(311, 175)
        Me._cmdRecord_4.Name = "_cmdRecord_4"
        Me._cmdRecord_4.Size = New System.Drawing.Size(75, 30)
        Me._cmdRecord_4.TabIndex = 32
        Me._cmdRecord_4.Text = "Cancel Edit"
        Me._cmdRecord_4.UseVisualStyleBackColor = False
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(73, 420)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(99, 43)
        Me._cmdCancelOK_1.TabIndex = 33
        Me._cmdCancelOK_1.Text = "OK"
        Me._cmdCancelOK_1.UseVisualStyleBackColor = False
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(245, 420)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(97, 43)
        Me._cmdCancelOK_0.TabIndex = 34
        Me._cmdCancelOK_0.Text = "Cancel"
        Me._cmdCancelOK_0.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtName)
        Me.GroupBox1.Controls.Add(Me._txtCoeff_1)
        Me.GroupBox1.Controls.Add(Me._lblUnit_3)
        Me.GroupBox1.Controls.Add(Me._txtCoeff_2)
        Me.GroupBox1.Controls.Add(Me._cmdRecord_4)
        Me.GroupBox1.Controls.Add(Me._cmdRecord_3)
        Me.GroupBox1.Controls.Add(Me._lblUnit_2)
        Me.GroupBox1.Controls.Add(Me._cmdRecord_2)
        Me.GroupBox1.Controls.Add(Me._txtCoeff_3)
        Me.GroupBox1.Controls.Add(Me._cmdRecord_1)
        Me.GroupBox1.Controls.Add(Me._lblUnit_1)
        Me.GroupBox1.Controls.Add(Me._cmdRecord_0)
        Me.GroupBox1.Controls.Add(Me._txtCoeff_4)
        Me.GroupBox1.Controls.Add(Me._lblUnit_0)
        Me.GroupBox1.Controls.Add(Me._lblDesc_1)
        Me.GroupBox1.Controls.Add(Me._lblDesc_4)
        Me.GroupBox1.Controls.Add(Me._lblDesc_2)
        Me.GroupBox1.Controls.Add(Me._lblDesc_3)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 189)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(408, 225)
        Me.GroupBox1.TabIndex = 35
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Empirical Kinetic Constants for:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lstCorrelations)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(408, 163)
        Me.GroupBox2.TabIndex = 36
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Select a Water Type:"
        '
        'frmFoulingWaterDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(425, 475)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(114, 177)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(441, 514)
        Me.Name = "frmFoulingWaterDatabase"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Water Fouling Correlation Database"
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCoeff, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents _cmdRecord_0 As Button
    Friend WithEvents _cmdRecord_1 As Button
    Friend WithEvents _cmdRecord_2 As Button
    Friend WithEvents _cmdRecord_3 As Button
    Friend WithEvents _cmdRecord_4 As Button
    Friend WithEvents _cmdCancelOK_1 As Button
    Friend WithEvents _cmdCancelOK_0 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
#End Region
End Class