<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSplash
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
	Public WithEvents _picLogos_0 As System.Windows.Forms.PictureBox
	Public WithEvents _lbldesc_0 As System.Windows.Forms.Label
	Public WithEvents _lbldesc_1 As System.Windows.Forms.Label
    Public WithEvents lblDisclaimer As System.Windows.Forms.Label
    Public WithEvents lblDisclaimerTitle As System.Windows.Forms.Label


    Public WithEvents cmdExit As System.Windows.Forms.Button
    Public WithEvents cmdButton2 As System.Windows.Forms.Button
    Public WithEvents cmdButton1 As System.Windows.Forms.Button
    Public WithEvents lblCompany As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVersionInfo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lbldesc As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents picLogos As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim _picLogos_1 As System.Windows.Forms.PictureBox
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSplash))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdButton2 = New System.Windows.Forms.Button()
        Me.cmdButton1 = New System.Windows.Forms.Button()
        Me.lblCompany = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVersionInfo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lbldesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lbldesc_0 = New System.Windows.Forms.Label()
        Me._lbldesc_1 = New System.Windows.Forms.Label()
        Me.picLogos = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(Me.components)
        Me._picLogos_0 = New System.Windows.Forms.PictureBox()
        Me.lblDisclaimerTitle = New System.Windows.Forms.Label()
        Me.lblDisclaimer = New System.Windows.Forms.Label()
        Me._picLogos_2 = New System.Windows.Forms.PictureBox()
        _picLogos_1 = New System.Windows.Forms.PictureBox()
        CType(_picLogos_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCompany, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVersionInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lbldesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._picLogos_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._picLogos_2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_picLogos_1
        '
        _picLogos_1.Image = CType(resources.GetObject("_picLogos_1.Image"), System.Drawing.Image)
        _picLogos_1.Location = New System.Drawing.Point(233, 46)
        _picLogos_1.Name = "_picLogos_1"
        _picLogos_1.Size = New System.Drawing.Size(145, 138)
        _picLogos_1.TabIndex = 21
        _picLogos_1.TabStop = False
        AddHandler _picLogos_1.Click, AddressOf Me._picLogos_1_Click
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(516, 338)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(87, 35)
        Me.cmdExit.TabIndex = 2
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdButton2
        '
        Me.cmdButton2.BackColor = System.Drawing.SystemColors.Control
        Me.cmdButton2.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdButton2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdButton2.Location = New System.Drawing.Point(626, 252)
        Me.cmdButton2.Name = "cmdButton2"
        Me.cmdButton2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdButton2.Size = New System.Drawing.Size(169, 35)
        Me.cmdButton2.TabIndex = 1
        Me.cmdButton2.Text = "I agree, never show again"
        Me.cmdButton2.UseVisualStyleBackColor = False
        '
        'cmdButton1
        '
        Me.cmdButton1.BackColor = System.Drawing.SystemColors.Control
        Me.cmdButton1.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdButton1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdButton1.Location = New System.Drawing.Point(6, 338)
        Me.cmdButton1.Name = "cmdButton1"
        Me.cmdButton1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdButton1.Size = New System.Drawing.Size(97, 35)
        Me.cmdButton1.TabIndex = 0
        Me.cmdButton1.Text = "&Continue"
        Me.cmdButton1.UseVisualStyleBackColor = False
        '
        '_lbldesc_0
        '
        Me._lbldesc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lbldesc_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lbldesc_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lbldesc_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbldesc.SetIndex(Me._lbldesc_0, CType(0, Short))
        Me._lbldesc_0.Location = New System.Drawing.Point(88, 48)
        Me._lbldesc_0.Name = "_lbldesc_0"
        Me._lbldesc_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lbldesc_0.Size = New System.Drawing.Size(137, 15)
        Me._lbldesc_0.TabIndex = 13
        Me._lbldesc_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lbldesc_1
        '
        Me._lbldesc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lbldesc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lbldesc_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lbldesc_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbldesc.SetIndex(Me._lbldesc_1, CType(1, Short))
        Me._lbldesc_1.Location = New System.Drawing.Point(88, 66)
        Me._lbldesc_1.Name = "_lbldesc_1"
        Me._lbldesc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lbldesc_1.Size = New System.Drawing.Size(137, 15)
        Me._lbldesc_1.TabIndex = 12
        Me._lbldesc_1.Text = " "
        Me._lbldesc_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_picLogos_0
        '
        Me._picLogos_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._picLogos_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._picLogos_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._picLogos_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._picLogos_0.Image = CType(resources.GetObject("_picLogos_0.Image"), System.Drawing.Image)
        Me.picLogos.SetIndex(Me._picLogos_0, CType(0, Short))
        Me._picLogos_0.Location = New System.Drawing.Point(24, 22)
        Me._picLogos_0.Name = "_picLogos_0"
        Me._picLogos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._picLogos_0.Size = New System.Drawing.Size(121, 117)
        Me._picLogos_0.TabIndex = 11
        Me._picLogos_0.TabStop = False
        '
        'lblDisclaimerTitle
        '
        Me.lblDisclaimerTitle.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDisclaimerTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDisclaimerTitle.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisclaimerTitle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDisclaimerTitle.Location = New System.Drawing.Point(4, 4)
        Me.lblDisclaimerTitle.Name = "lblDisclaimerTitle"
        Me.lblDisclaimerTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDisclaimerTitle.Size = New System.Drawing.Size(545, 35)
        Me.lblDisclaimerTitle.TabIndex = 7
        Me.lblDisclaimerTitle.Text = "Disclaimer"
        Me.lblDisclaimerTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblDisclaimer
        '
        Me.lblDisclaimer.BackColor = System.Drawing.Color.Transparent
        Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDisclaimer.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisclaimer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDisclaimer.Location = New System.Drawing.Point(6, 6)
        Me.lblDisclaimer.Name = "lblDisclaimer"
        Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDisclaimer.Size = New System.Drawing.Size(523, 193)
        Me.lblDisclaimer.TabIndex = 9
        Me.lblDisclaimer.Text = "lblDisclaimer"
        '
        '_picLogos_2
        '
        Me._picLogos_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._picLogos_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._picLogos_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._picLogos_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._picLogos_2.Image = CType(resources.GetObject("_picLogos_2.Image"), System.Drawing.Image)
        Me._picLogos_2.Location = New System.Drawing.Point(201, 230)
        Me._picLogos_2.Name = "_picLogos_2"
        Me._picLogos_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._picLogos_2.Size = New System.Drawing.Size(218, 48)
        Me._picLogos_2.TabIndex = 22
        Me._picLogos_2.TabStop = False
        '
        'frmSplash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(607, 382)
        Me.Controls.Add(Me._picLogos_2)
        Me.Controls.Add(_picLogos_1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdButton2)
        Me.Controls.Add(Me.cmdButton1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(83, 105)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSplash"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Program Information"
        CType(_picLogos_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCompany, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVersionInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lbldesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._picLogos_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._picLogos_2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public WithEvents _picLogos_2 As PictureBox
#End Region
End Class