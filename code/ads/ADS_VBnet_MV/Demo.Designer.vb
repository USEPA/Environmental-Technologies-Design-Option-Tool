<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDemo
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
	Public WithEvents cmdButton1 As System.Windows.Forms.Button
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents lblDisclaimer As System.Windows.Forms.Label
    Public WithEvents SSPanel3 As AxThreed.AxSSPanel
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDemo))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdButton1 = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.SSPanel3 = New AxThreed.AxSSPanel()
        Me.lblDisclaimer = New System.Windows.Forms.Label()
        CType(Me.SSPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdButton1
        '
        Me.cmdButton1.BackColor = System.Drawing.SystemColors.Control
        Me.cmdButton1.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdButton1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdButton1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdButton1.Location = New System.Drawing.Point(8, 220)
        Me.cmdButton1.Name = "cmdButton1"
        Me.cmdButton1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdButton1.Size = New System.Drawing.Size(97, 35)
        Me.cmdButton1.TabIndex = 0
        Me.cmdButton1.Text = "&Continue"
        Me.cmdButton1.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(460, 220)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(87, 35)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'SSPanel3
        '
        Me.SSPanel3.Controls.Add(Me.lblDisclaimer)
        Me.SSPanel3.Location = New System.Drawing.Point(8, 8)
        Me.SSPanel3.Name = "SSPanel3"
        Me.SSPanel3.OcxState = CType(resources.GetObject("SSPanel3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSPanel3.Size = New System.Drawing.Size(539, 205)
        Me.SSPanel3.TabIndex = 2
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
        Me.lblDisclaimer.TabIndex = 3
        Me.lblDisclaimer.Text = "lblDisclaimer"
        '
        'frmDemo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(555, 265)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdButton1)
        Me.Controls.Add(Me.cmdExit)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(182, 156)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDemo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Demo Version"
        CType(Me.SSPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSPanel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class