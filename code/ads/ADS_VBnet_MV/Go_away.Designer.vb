<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmGoAway
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
	Public WithEvents chkDisplay As System.Windows.Forms.CheckBox
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmGoAway))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkDisplay = New System.Windows.Forms.CheckBox
		Me.Picture1 = New System.Windows.Forms.PictureBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.cmdOK = New System.Windows.Forms.Button
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.Color.White
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "{Caller-Defined Caption}"
		Me.ClientSize = New System.Drawing.Size(486, 214)
		Me.Location = New System.Drawing.Point(141, 263)
		Me.ControlBox = False
		Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ForeColor = System.Drawing.SystemColors.WindowText
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmGoAway"
		Me.chkDisplay.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.chkDisplay.BackColor = System.Drawing.SystemColors.Window
		Me.chkDisplay.Text = "{Caller-Defined Text}"
		Me.chkDisplay.ForeColor = System.Drawing.SystemColors.WindowText
		Me.chkDisplay.Size = New System.Drawing.Size(361, 25)
		Me.chkDisplay.Location = New System.Drawing.Point(16, 180)
		Me.chkDisplay.TabIndex = 4
		Me.chkDisplay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDisplay.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDisplay.CausesValidation = True
		Me.chkDisplay.Enabled = True
		Me.chkDisplay.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDisplay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDisplay.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDisplay.TabStop = True
		Me.chkDisplay.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDisplay.Visible = True
		Me.chkDisplay.Name = "chkDisplay"
		Me.Picture1.BackColor = System.Drawing.SystemColors.Window
		Me.Picture1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Picture1.Size = New System.Drawing.Size(57, 29)
		Me.Picture1.Location = New System.Drawing.Point(464, -4)
		Me.Picture1.TabIndex = 3
		Me.Picture1.Visible = False
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Picture1.Name = "Picture1"
		Me.Frame1.BackColor = System.Drawing.Color.White
		Me.Frame1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Frame1.Size = New System.Drawing.Size(461, 161)
		Me.Frame1.Location = New System.Drawing.Point(12, 8)
		Me.Frame1.TabIndex = 1
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Label1.BackColor = System.Drawing.Color.White
		Me.Label1.Text = "{Caller-Defined Text}"
		Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Label1.Size = New System.Drawing.Size(441, 145)
		Me.Label1.Location = New System.Drawing.Point(8, 12)
		Me.Label1.TabIndex = 2
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "&OK"
		Me.cmdOK.Size = New System.Drawing.Size(81, 33)
		Me.cmdOK.Location = New System.Drawing.Point(392, 176)
		Me.cmdOK.TabIndex = 0
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.Controls.Add(chkDisplay)
		Me.Controls.Add(Picture1)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdOK)
		Me.Frame1.Controls.Add(Label1)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class