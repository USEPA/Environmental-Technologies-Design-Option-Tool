<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFluidProps
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
    Public WithEvents _chkCorr_0 As AxThreed.AxSSCheck
    Public WithEvents _chkCorr_1 As AxThreed.AxSSCheck


    Public WithEvents SSFrame2 As AxThreed.AxSSFrame

    Friend WithEvents _lblunit_0 As System.Windows.Forms.Label
    Friend WithEvents _lblunit_1 As System.Windows.Forms.Label
    Friend WithEvents _txtWater_0 As System.Windows.Forms.TextBox
    Friend WithEvents _txtWater_1 As System.Windows.Forms.TextBox
    ' Friend WithEvents Label1 As Label
    ' Friend WithEvents Label2 As Label
    ' Friend WithEvents TextBox1 As TextBox
    ' Friend WithEvents TextBox2 As TextBox
    Public WithEvents SSFrame3 As AxThreed.AxSSFrame
    Public WithEvents SSFrame1 As AxThreed.AxSSFrame
    Public WithEvents _cmdCancelOK_1 As AxThreed.AxSSCommand
    Public WithEvents _cmdCancelOK_0 As AxThreed.AxSSCommand
    '   Public WithEvents chkCorr As SSCheckArray
    '   Public WithEvents cmdCancelOK As SSCommandArray
    Public WithEvents lblUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtWater As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFluidProps))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblunit_0 = New System.Windows.Forms.Label()
        Me._lblunit_1 = New System.Windows.Forms.Label()
        Me.txtWater = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me._txtWater_0 = New System.Windows.Forms.TextBox()
        Me._txtWater_1 = New System.Windows.Forms.TextBox()
        Me.SSFrame1 = New AxThreed.AxSSFrame()
        Me._cmdCancelOK_1 = New AxThreed.AxSSCommand()
        Me._cmdCancelOK_0 = New AxThreed.AxSSCommand()
        Me._chkCorr_0 = New AxThreed.AxSSCheck()
        Me._chkCorr_1 = New AxThreed.AxSSCheck()
        Me.SSFrame2 = New AxThreed.AxSSFrame()
        Me.SSFrame3 = New AxThreed.AxSSFrame()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWater, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdCancelOK_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdCancelOK_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkCorr_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkCorr_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSFrame2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSFrame3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(365, 171)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(97, 22)
        Me.Command4.TabIndex = 11
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
        Me.Picture1.Location = New System.Drawing.Point(390, 30)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 12
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        '_lblunit_0
        '
        Me._lblunit_0.AutoSize = True
        Me.lblUnit.SetIndex(Me._lblunit_0, CType(0, Short))
        Me._lblunit_0.Location = New System.Drawing.Point(101, 198)
        Me._lblunit_0.Name = "_lblunit_0"
        Me._lblunit_0.Size = New System.Drawing.Size(64, 16)
        Me._lblunit_0.TabIndex = 13
        Me._lblunit_0.Text = "lblUnit(0)"
        '
        '_lblunit_1
        '
        Me._lblunit_1.AutoSize = True
        Me.lblUnit.SetIndex(Me._lblunit_1, CType(1, Short))
        Me._lblunit_1.Location = New System.Drawing.Point(104, 245)
        Me._lblunit_1.Name = "_lblunit_1"
        Me._lblunit_1.Size = New System.Drawing.Size(64, 16)
        Me._lblunit_1.TabIndex = 14
        Me._lblunit_1.Text = "lblUnit(1)"
        '
        'txtWater
        '
        '
        '_txtWater_0
        '
        Me.txtWater.SetIndex(Me._txtWater_0, CType(0, Short))
        Me._txtWater_0.Location = New System.Drawing.Point(199, 198)
        Me._txtWater_0.Name = "_txtWater_0"
        Me._txtWater_0.Size = New System.Drawing.Size(100, 23)
        Me._txtWater_0.TabIndex = 15
        Me._txtWater_0.Text = "txtWater(0)"
        '
        '_txtWater_1
        '
        Me.txtWater.SetIndex(Me._txtWater_1, CType(1, Short))
        Me._txtWater_1.Location = New System.Drawing.Point(199, 242)
        Me._txtWater_1.Name = "_txtWater_1"
        Me._txtWater_1.Size = New System.Drawing.Size(100, 23)
        Me._txtWater_1.TabIndex = 16
        Me._txtWater_1.Text = "txtWater(1)"
        '
        'SSFrame1
        '
        Me.SSFrame1.Location = New System.Drawing.Point(168, 12)
        Me.SSFrame1.Name = "SSFrame1"
        Me.SSFrame1.OcxState = CType(resources.GetObject("SSFrame1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame1.Size = New System.Drawing.Size(166, 101)
        Me.SSFrame1.TabIndex = 4
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(68, 148)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.OcxState = CType(resources.GetObject("_cmdCancelOK_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(100, 32)
        Me._cmdCancelOK_1.TabIndex = 9
        Me._cmdCancelOK_1.TabStop = False
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(199, 148)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.OcxState = CType(resources.GetObject("_cmdCancelOK_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(100, 32)
        Me._cmdCancelOK_0.TabIndex = 10
        Me._cmdCancelOK_0.TabStop = False
        '
        '_chkCorr_0
        '
        Me._chkCorr_0.Location = New System.Drawing.Point(37, 12)
        Me._chkCorr_0.Name = "_chkCorr_0"
        Me._chkCorr_0.OcxState = CType(resources.GetObject("_chkCorr_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkCorr_0.Size = New System.Drawing.Size(100, 38)
        Me._chkCorr_0.TabIndex = 0
        '
        '_chkCorr_1
        '
        Me._chkCorr_1.Location = New System.Drawing.Point(37, 76)
        Me._chkCorr_1.Name = "_chkCorr_1"
        Me._chkCorr_1.OcxState = CType(resources.GetObject("_chkCorr_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkCorr_1.Size = New System.Drawing.Size(100, 30)
        Me._chkCorr_1.TabIndex = 1
        '
        'SSFrame2
        '
        Me.SSFrame2.Location = New System.Drawing.Point(8, 18)
        Me.SSFrame2.Name = "SSFrame2"
        Me.SSFrame2.OcxState = CType(resources.GetObject("SSFrame2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame2.Size = New System.Drawing.Size(127, 67)
        Me.SSFrame2.TabIndex = 5
        '
        'SSFrame3
        '
        Me.SSFrame3.Location = New System.Drawing.Point(134, 18)
        Me.SSFrame3.Name = "SSFrame3"
        Me.SSFrame3.OcxState = CType(resources.GetObject("SSFrame3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame3.Size = New System.Drawing.Size(163, 67)
        Me.SSFrame3.TabIndex = 6
        '
        'frmFluidProps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(491, 316)
        Me.ControlBox = False
        Me.Controls.Add(Me._txtWater_0)
        Me.Controls.Add(Me._txtWater_1)
        Me.Controls.Add(Me._lblunit_0)
        Me.Controls.Add(Me._lblunit_1)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.SSFrame1)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Controls.Add(Me._chkCorr_0)
        Me.Controls.Add(Me._chkCorr_1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(282, 232)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFluidProps"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "{Fluid} Properties"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWater, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdCancelOK_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdCancelOK_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkCorr_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkCorr_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSFrame2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSFrame3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub




#End Region
End Class