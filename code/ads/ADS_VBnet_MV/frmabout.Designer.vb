<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAbout2
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
    Public WithEvents _pnl_title_3 As AxThreed.AxSSPanel
    Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents _pnl_titleX_5 As System.Windows.Forms.PictureBox
    '   Public WithEvents pnl_title As SSPanelArray
    Public WithEvents pnl_titleX As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout2))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me._pnl_title_3 = New AxThreed.AxSSPanel
        Me.cmdOK = New System.Windows.Forms.Button
        Me._pnl_titleX_5 = New System.Windows.Forms.PictureBox
        '     Me.pnl_title = New SSPanelArray(components)
        Me.pnl_titleX = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        '        CType(Me.pnl_title, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnl_titleX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Text = "Technical Support Provided By:"
        Me.ClientSize = New System.Drawing.Size(355, 198)
        Me.Location = New System.Drawing.Point(153, 453)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ControlBox = True
        Me.Enabled = True
        Me.KeyPreview = False
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmAbout2"
        Me._pnl_title_3.Size = New System.Drawing.Size(347, 147)
        Me._pnl_title_3.Location = New System.Drawing.Point(4, 16)
        Me._pnl_title_3.TabIndex = 2
        Me._pnl_title_3.Caption = "pnl_title(3)"
        Me._pnl_title_3.Name = "_pnl_title_3"
        Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.CancelButton = Me.cmdOK
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.Size = New System.Drawing.Size(91, 23)
        Me.cmdOK.Location = New System.Drawing.Point(260, 170)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.CausesValidation = True
        Me.cmdOK.Enabled = True
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.TabStop = True
        Me.cmdOK.Name = "cmdOK"
        Me._pnl_titleX_5.BackColor = System.Drawing.SystemColors.Window
        Me._pnl_titleX_5.ForeColor = System.Drawing.Color.FromARGB(64, 0, 64)
        Me._pnl_titleX_5.Size = New System.Drawing.Size(347, 49)
        Me._pnl_titleX_5.Location = New System.Drawing.Point(68, 220)
        Me._pnl_titleX_5.TabIndex = 0
        Me._pnl_titleX_5.Visible = False
        Me._pnl_titleX_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._pnl_titleX_5.Dock = System.Windows.Forms.DockStyle.None
        Me._pnl_titleX_5.CausesValidation = True
        Me._pnl_titleX_5.Enabled = True
        Me._pnl_titleX_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._pnl_titleX_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._pnl_titleX_5.TabStop = True
        Me._pnl_titleX_5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
        Me._pnl_titleX_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._pnl_titleX_5.Name = "_pnl_titleX_5"
        Me.Controls.Add(_pnl_title_3)
        Me.Controls.Add(cmdOK)
        Me.Controls.Add(_pnl_titleX_5)
        '     Me.pnl_title.SetIndex(_pnl_title_3, CType(3, Short))
        '    Me.pnl_titleX.SetIndex(_pnl_titleX_5, CType(5, Short))
        CType(Me.pnl_titleX, System.ComponentModel.ISupportInitialize).EndInit()
        '   CType(Me.pnl_title, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
#End Region 
End Class