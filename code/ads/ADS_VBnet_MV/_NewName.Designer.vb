<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewName
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
	Public WithEvents txtdata As System.Windows.Forms.TextBox
    Public WithEvents _Button_0 As AxThreed.AxSSCommand
    Public WithEvents _Button_1 As AxThreed.AxSSCommand
    Public WithEvents lblInstructions As System.Windows.Forms.Label
    '   Public WithEvents Button As SSCommandArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNewName))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtdata = New System.Windows.Forms.TextBox()
        Me._Button_0 = New AxThreed.AxSSCommand()
        Me._Button_1 = New AxThreed.AxSSCommand()
        Me.lblInstructions = New System.Windows.Forms.Label()
        CType(Me._Button_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._Button_1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtdata
        '
        Me.txtdata.AcceptsReturn = True
        Me.txtdata.BackColor = System.Drawing.SystemColors.Window
        Me.txtdata.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtdata.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdata.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtdata.Location = New System.Drawing.Point(6, 50)
        Me.txtdata.MaxLength = 0
        Me.txtdata.Name = "txtdata"
        Me.txtdata.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtdata.Size = New System.Drawing.Size(367, 26)
        Me.txtdata.TabIndex = 0
        Me.txtdata.Text = "txtdata"
        '
        '_Button_0
        '
        Me._Button_0.Location = New System.Drawing.Point(6, 80)
        Me._Button_0.Name = "_Button_0"
        Me._Button_0.OcxState = CType(resources.GetObject("_Button_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._Button_0.Size = New System.Drawing.Size(100, 50)
        Me._Button_0.TabIndex = 2
        Me._Button_0.TabStop = False
        '
        '_Button_1
        '
        Me._Button_1.Location = New System.Drawing.Point(273, 80)
        Me._Button_1.Name = "_Button_1"
        Me._Button_1.OcxState = CType(resources.GetObject("_Button_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._Button_1.Size = New System.Drawing.Size(100, 50)
        Me._Button_1.TabIndex = 3
        Me._Button_1.TabStop = False
        '
        'lblInstructions
        '
        Me.lblInstructions.BackColor = System.Drawing.SystemColors.Control
        Me.lblInstructions.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInstructions.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInstructions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInstructions.Location = New System.Drawing.Point(6, 4)
        Me.lblInstructions.Name = "lblInstructions"
        Me.lblInstructions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInstructions.Size = New System.Drawing.Size(367, 41)
        Me.lblInstructions.TabIndex = 1
        Me.lblInstructions.Text = "lblInstructions"
        '
        'frmNewName
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(485, 154)
        Me.ControlBox = False
        Me.Controls.Add(Me.txtdata)
        Me.Controls.Add(Me._Button_0)
        Me.Controls.Add(Me._Button_1)
        Me.Controls.Add(Me.lblInstructions)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(117, 492)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNewName"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "{use_title}"
        CType(Me._Button_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._Button_1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class