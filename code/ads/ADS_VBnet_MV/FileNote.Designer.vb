<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFileNote
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
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents txtData As System.Windows.Forms.TextBox
    Public WithEvents _cmdButton_0 As AxThreed.AxSSCommand
    Public WithEvents _cmdButton_1 As AxThreed.AxSSCommand
    Public WithEvents _cmdButton_2 As AxThreed.AxSSCommand
    Public WithEvents _cmdButton_3 As AxThreed.AxSSCommand
    Public WithEvents _cmdButton_4 As AxThreed.AxSSCommand
    Public WithEvents _cmdButton_5 As AxThreed.AxSSCommand
    Public cmdButton(6) As AxThreed.AxSSCommand    'Shang add control array
    Public WithEvents lblInstructions As System.Windows.Forms.Label
    '  Public WithEvents cmdButton As SSCommandArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFileNote))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.txtData = New System.Windows.Forms.TextBox()
        Me._cmdButton_0 = New AxThreed.AxSSCommand()
        Me._cmdButton_1 = New AxThreed.AxSSCommand()
        Me._cmdButton_2 = New AxThreed.AxSSCommand()
        Me._cmdButton_3 = New AxThreed.AxSSCommand()
        Me._cmdButton_4 = New AxThreed.AxSSCommand()
        Me._cmdButton_5 = New AxThreed.AxSSCommand()
        Me.lblInstructions = New System.Windows.Forms.Label()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdButton_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdButton_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdButton_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdButton_3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdButton_4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdButton_5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(200, 246)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(143, 22)
        Me.Command4.TabIndex = 8
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
        Me.Picture1.Location = New System.Drawing.Point(604, 211)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 9
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'txtData
        '
        Me.txtData.AcceptsReturn = True
        Me.txtData.BackColor = System.Drawing.SystemColors.Window
        Me.txtData.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtData.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtData.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.Location = New System.Drawing.Point(6, 34)
        Me.txtData.MaxLength = 500
        Me.txtData.Multiline = True
        Me.txtData.Name = "txtData"
        Me.txtData.ReadOnly = True
        Me.txtData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtData.Size = New System.Drawing.Size(572, 207)
        Me.txtData.TabIndex = 0
        Me.txtData.Text = "txtData"
        '
        '_cmdButton_0
        '
        Me._cmdButton_0.Location = New System.Drawing.Point(6, 272)
        Me._cmdButton_0.Name = "_cmdButton_0"
        Me._cmdButton_0.OcxState = CType(resources.GetObject("_cmdButton_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdButton_0.Size = New System.Drawing.Size(81, 50)
        Me._cmdButton_0.TabIndex = 2
        Me._cmdButton_0.TabStop = False
        Me.cmdButton(0) = _cmdButton_0
        AddHandler _cmdButton_0.ClickEvent, AddressOf cmdButton_Click
        '
        '_cmdButton_1
        '
        Me._cmdButton_1.Location = New System.Drawing.Point(84, 272)
        Me._cmdButton_1.Name = "_cmdButton_1"
        Me._cmdButton_1.OcxState = CType(resources.GetObject("_cmdButton_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdButton_1.Size = New System.Drawing.Size(84, 50)
        Me._cmdButton_1.TabIndex = 3
        Me._cmdButton_1.TabStop = False
        Me.cmdButton(1) = _cmdButton_1
        AddHandler _cmdButton_1.ClickEvent, AddressOf cmdbutton_click
        '
        '_cmdButton_2
        '
        Me._cmdButton_2.Location = New System.Drawing.Point(167, 272)
        Me._cmdButton_2.Name = "_cmdButton_2"
        Me._cmdButton_2.OcxState = CType(resources.GetObject("_cmdButton_2.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdButton_2.Size = New System.Drawing.Size(82, 50)
        Me._cmdButton_2.TabIndex = 4
        Me._cmdButton_2.TabStop = False
        Me.cmdButton(2) = _cmdButton_2
        AddHandler _cmdButton_2.ClickEvent, AddressOf cmdbutton_click
        '
        '_cmdButton_3
        '
        Me._cmdButton_3.Location = New System.Drawing.Point(406, 272)
        Me._cmdButton_3.Name = "_cmdButton_3"
        Me._cmdButton_3.OcxState = CType(resources.GetObject("_cmdButton_3.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdButton_3.Size = New System.Drawing.Size(77, 50)
        Me._cmdButton_3.TabIndex = 5
        Me._cmdButton_3.TabStop = False
        Me.cmdButton(3) = _cmdButton_3
        AddHandler _cmdButton_3.ClickEvent, AddressOf cmdbutton_click
        '
        '_cmdButton_4
        '
        Me._cmdButton_4.Location = New System.Drawing.Point(478, 272)
        Me._cmdButton_4.Name = "_cmdButton_4"
        Me._cmdButton_4.OcxState = CType(resources.GetObject("_cmdButton_4.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdButton_4.Size = New System.Drawing.Size(100, 50)
        Me._cmdButton_4.TabIndex = 6
        Me._cmdButton_4.TabStop = False
        Me.cmdButton(4) = _cmdButton_4
        AddHandler _cmdButton_4.ClickEvent, AddressOf cmdbutton_click
        '
        '_cmdButton_5
        '
        Me._cmdButton_5.Location = New System.Drawing.Point(273, 271)
        Me._cmdButton_5.Name = "_cmdButton_5"
        Me._cmdButton_5.OcxState = CType(resources.GetObject("_cmdButton_5.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdButton_5.Size = New System.Drawing.Size(133, 50)
        Me._cmdButton_5.TabIndex = 7
        Me._cmdButton_5.TabStop = False
        Me.cmdButton(5) = _cmdButton_5
        AddHandler _cmdButton_5.ClickEvent, AddressOf cmdbutton_click
        '
        'lblInstructions
        '
        Me.lblInstructions.BackColor = System.Drawing.SystemColors.Control
        Me.lblInstructions.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInstructions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInstructions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInstructions.Location = New System.Drawing.Point(6, 2)
        Me.lblInstructions.Name = "lblInstructions"
        Me.lblInstructions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInstructions.Size = New System.Drawing.Size(397, 27)
        Me.lblInstructions.TabIndex = 1
        Me.lblInstructions.Text = "You may enter up to 500 characters of text.  Line breaks are acceptable."
        '
        'frmFileNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(598, 333)
        Me.ControlBox = False
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.txtData)
        Me.Controls.Add(Me._cmdButton_0)
        Me.Controls.Add(Me._cmdButton_1)
        Me.Controls.Add(Me._cmdButton_2)
        Me.Controls.Add(Me._cmdButton_3)
        Me.Controls.Add(Me._cmdButton_4)
        Me.Controls.Add(Me._cmdButton_5)
        Me.Controls.Add(Me.lblInstructions)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(152, 304)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFileNote"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "File Note"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdButton_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdButton_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdButton_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdButton_3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdButton_4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdButton_5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class