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
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        _picLogos_1 = New System.Windows.Forms.PictureBox()
        CType(_picLogos_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCompany, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVersionInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lbldesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._picLogos_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._picLogos_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_picLogos_1
        '
        _picLogos_1.Image = CType(resources.GetObject("_picLogos_1.Image"), System.Drawing.Image)
        _picLogos_1.Location = New System.Drawing.Point(28, 200)
        _picLogos_1.Name = "_picLogos_1"
        _picLogos_1.Size = New System.Drawing.Size(133, 110)
        _picLogos_1.TabIndex = 21
        _picLogos_1.TabStop = False
        AddHandler _picLogos_1.Click, AddressOf Me._picLogos_1_Click
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(483, 491)
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
        Me.cmdButton1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdButton1.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdButton1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdButton1.Location = New System.Drawing.Point(33, 491)
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
        Me._picLogos_2.Location = New System.Drawing.Point(28, 360)
        Me._picLogos_2.Name = "_picLogos_2"
        Me._picLogos_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._picLogos_2.Size = New System.Drawing.Size(217, 48)
        Me._picLogos_2.TabIndex = 22
        Me._picLogos_2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(28, 54)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(102, 102)
        Me.PictureBox1.TabIndex = 23
        Me.PictureBox1.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Font = New System.Drawing.Font("Impact", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(119, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(386, 26)
        Me.TextBox1.TabIndex = 24
        Me.TextBox1.Text = "Adsorption Design Software (AdDesignS™)"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(182, 54)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(255, 365)
        Me.TextBox2.TabIndex = 25
        Me.TextBox2.Text = resources.GetString("TextBox2.Text")
        '
        'TextBox3
        '
        Me.TextBox3.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(443, 79)
        Me.TextBox3.Multiline = True
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(137, 111)
        Me.TextBox3.TabIndex = 26
        Me.TextBox3.Text = "Open-Source Edition" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Michael Verma" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Feng Shang" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.Label1.Location = New System.Drawing.Point(72, 442)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(470, 16)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "http://github.com/USEPA/Environmental-Technologies-Design-Option-Tool"
        '
        'TextBox4
        '
        Me.TextBox4.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox4.Location = New System.Drawing.Point(443, 234)
        Me.TextBox4.Multiline = True
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(137, 164)
        Me.TextBox4.TabIndex = 28
        Me.TextBox4.Text = "Copyright 1994-2005" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "David R. Hokanson" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "David W. Hand" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "John C. Crittenden" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Tony" &
    " N. Rogers" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Eric J. Oman"
        '
        'TextBox5
        '
        Me.TextBox5.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox5.Location = New System.Drawing.Point(443, 54)
        Me.TextBox5.Multiline = True
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(137, 19)
        Me.TextBox5.TabIndex = 29
        Me.TextBox5.Text = "Version 1.0.50"
        '
        'TextBox6
        '
        Me.TextBox6.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox6.Location = New System.Drawing.Point(443, 209)
        Me.TextBox6.Multiline = True
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(137, 19)
        Me.TextBox6.TabIndex = 30
        Me.TextBox6.Text = "Version 1.0"
        '
        'frmSplash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(607, 538)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(_picLogos_1)
        Me.Controls.Add(Me._picLogos_2)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdButton2)
        Me.Controls.Add(Me.cmdButton1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Location = New System.Drawing.Point(83, 105)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(623, 577)
        Me.Name = "frmSplash"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AdDesignS"
        CType(_picLogos_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCompany, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVersionInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lbldesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._picLogos_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._picLogos_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public WithEvents _picLogos_2 As PictureBox
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents TextBox6 As TextBox
#End Region
End Class